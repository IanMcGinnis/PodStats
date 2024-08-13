import asyncio
import json
import math
import os
from datetime import datetime
from difflib import get_close_matches
from functools import partial

import discord
import gspread
import pandas as pd
from dotenv import load_dotenv
from discord import app_commands
from discord.ui import Button, View
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

# Discord set up
load_dotenv()
discord_key = os.getenv('DiscordToken')
intents = discord.Intents.all()
bot = discord.Client(intents=intents)
tree = app_commands.CommandTree(bot)

# Google set up
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('PodStatsAuth.json', scope)
client = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)


# Files to store information
channels_file = 'channels.json'
used_commands_file = 'used_commands.json'
active_stats_file = 'active_stats.json'

active_game = {
}

def load_channels():
    if os.path.exists(channels_file):
        with open(channels_file, 'r') as f:
            return json.load(f)
    return {}
def save_channels(channels):
    with open(channels_file, 'w') as f:
        json.dump(channels, f, indent=4)

def is_correct_channel(interaction: discord.Interaction):
    return str(interaction.guild.id) in channels and channels[str(interaction.guild.id)] == interaction.channel.id
def is_active_game(interaction: discord.Interaction):
    return interaction.guild.id in active_game

def load_active_stats():
    if os.path.exists(active_stats_file):
        with open(active_stats_file, 'r') as f:
            return json.load(f)
    return {}
def save_active_stats(active_stats):
    with open(active_stats_file, 'w') as f:
        json.dump(active_stats, f, indent=4)

channels = load_channels()
active_sheets = load_active_stats()

class MyView(View):
    def __init__(self, players, commanders):
        super().__init__()

        self.buttons = {i: [] for i in range(len(players))}
        self.col0, self.col1, self.col2 = False, False, False
        self.playerOut, self.playerWon, self.playerBlood = '', '', ''

        # making each button
        firstOutButtons = [Button(label=f'first out {players[i]}', style=discord.ButtonStyle.red) for i in range(len(players))]
        winnerButtons = [Button(label=f'won {players[i]}', style=discord.ButtonStyle.green) for i in range(len(players))]
        firstBloodButtons = [Button(label=f'first blood {players[i]}', style=discord.ButtonStyle.blurple) for i in range(len(players))]

        #putting all buttons in the array in proper order
        for i in range(len(players)):
            # assigning each button the current row
            firstOutButtons[i].row = i
            winnerButtons[i].row = i
            firstBloodButtons[i].row = i

            #assigning each button a function
            firstOutButtons[i].callback = partial(self.button_callback, column=0, player= players[i])
            winnerButtons[i].callback = partial(self.button_callback, column=1, player= players[i])
            firstBloodButtons[i].callback = partial(self.button_callback, column=2, player= players[i])

            #adding the buttons to the button list
            self.add_item(firstOutButtons[i])
            self.add_item(winnerButtons[i])
            self.add_item(firstBloodButtons[i])

            # storing buttons in dictionary by column
            self.buttons[i].extend([firstOutButtons[i], winnerButtons[i], firstBloodButtons[i]])
    async def button_callback(self, interaction: discord.Interaction, column, player):
        tempcol0, tempcol1, tempcol2 = False, False, False
        for row in range(len(self.buttons)):
            self.buttons[row][column].disabled = True

            if column == 0 and not self.col0:
                self.col0, tempcol0 = True, True
                self.playerOut = player

            if column == 1 and not self.col1:
                self.col1, tempcol1 = True, True
                self.playerWon = player

            if column == 2 and not self.col2:
                self.col2, tempcol2 = True, True
                self.playerBlood = player

        if tempcol0 and self.col0:
            await interaction.response.edit_message(content=f'You selected: {player} for first out', view=self)
        if tempcol1 and self.col1:
            await interaction.response.edit_message(content=f'You selected: {player} for winning the game', view=self)
        if tempcol2 and self.col2:
            await interaction.response.edit_message(content=f'You selected: {player} for first blood', view=self)

        if self.col0 and self.col1 and self.col2:
            await interaction.followup.send(f'Your selection: \n'
                                            f'player that got out first:{self.playerOut},\n'
                                            f'player that won {self.playerWon},\n'
                                            f'player that died first: {self.playerBlood}\n\n'
                                            f'Finishing game!')

            sheetID = active_sheets[str(interaction.guild.id)]

            Title, Tables, RawData, Validations = get_sheets(sheetID)
            spreadsheet = client.open(Title)

            finish_game_stats(interaction, players= [self.playerOut, self.playerWon, self.playerBlood], RawData=spreadsheet.worksheet(RawData))
        # Update the label to show it has been clicked
        #await interaction.message.edit(content=interaction.message.content, view=self)

def finish_game_stats(interaction: discord.Interaction, players, RawData):
    gameNumber = RawData.col_values(2)  # Get all values in the column
    lastGameRow = len(gameNumber) - 3  # The last filled row is the length of the list
    newGameNumber = math.ceil((lastGameRow) / 4)
    print(f'finishing game {active_game[interaction.guild.id]} . . .')

    #grab todays date
    now = datetime.now()
    # Format the date as MM/DD/YYYY
    formattedDate = now.strftime('%m/%d/%Y')

    startRow = lastGameRow  # Ensure start_row is at least 1
    endRow = lastGameRow + 3

    # Fetch the last 4 rows in columns D, E, and F
    rangeString = f'B{startRow}:F{endRow}'
    last4Rows = RawData.get(rangeString)


    for row in last4Rows:
        if row[0] in players:
            for i, stat in enumerate(players):
                if stat == row[0]:
                    row[i+2] = 1

        # Update the spreadsheet with the modified rows
    for i, updatedRow in enumerate(last4Rows):
        cell_range = f'B{startRow + i}:F{startRow + i}'
        RawData.update([updatedRow], cell_range)
    del active_game[interaction.guild.id]
    print(f'game {newGameNumber} edited on {formattedDate}')

def game_to_sheet(players, commanders, RawData):
    gameNumber = RawData.col_values(1)  # Get all values in the column
    lastGameRow = len(gameNumber)  # The last filled row is the length of the list
    newGameNumber = math.ceil((lastGameRow) / 4)

    #grab todays date
    now = datetime.now()
    # Format the date as MM/DD/YYYY
    formattedDate = now.strftime('%m/%d/%Y')

    for i in range(len(players)):
        row = [newGameNumber, players[i].capitalize(), commanders[i].capitalize(), 0, 0, 0, formattedDate]
        RawData.append_row(row)

    print(f'Added game {newGameNumber} on {formattedDate}')

def list_files():
    results = drive_service.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet'",
        pageSize=10,
        fields="nextPageToken, files(id, name)"
    ).execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
        return None
    else:
        print('Files:')
        for item in items:
            print(f"{item['name']} ({item['id']})")
        return items

def get_sheets(spreadsheetID):
    """Retrieve the names of all sheets in the spreadsheet."""
    sheetMetadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetID).execute()
    title = sheetMetadata.get('properties', {}).get('title')
    sheets = sheetMetadata.get('sheets', [])
    sheetNames = [sheet['properties']['title'] for sheet in sheets]

    return title, sheetNames[0], sheetNames[1], sheetNames[2]
def data_refresh(Valid):
    # Get all values from the sheet
    data = Valid.get_all_values()

    # Print the data
    listOfCommanders = []
    listOfPlayers = []

    for row in data:
        if row[0] != '':
            listOfCommanders.append(row[0].capitalize())
        if row[3] != '':
            listOfPlayers.append(row[3].capitalize())
    #remove column title
    listOfPlayers = listOfPlayers[1:]
    listOfCommanders = listOfCommanders[1:]

    #alphabatize
    listOfPlayers.sort()
    listOfCommanders.sort()

    print('Grabbed player and commander names from spreadsheet . . .')
    return listOfPlayers, listOfCommanders

def correct_name(inputNames, knownNames):
    fixedNames = []

    for name in inputNames:
        name = name.capitalize()
        matches = get_close_matches(word=name, possibilities=knownNames, n=1, cutoff=0.4)

        if matches:
            fixedNames.append(matches[0])
        else:
            fixedNames.append(name)
    return fixedNames

@bot.event
async def on_guild_join(guild):
    # When the bot joins a new guild, send a 'hi' message to the first available text channel
    print(f'joined {guild.name}')
    for channel in guild.text_channels:
        if channel.permissions_for(guild.me).send_messages:
            await channel.send('To set me up, please do /setup in the channel I will be in!')
            break

@tree.command(name='setup', description='Sets up the bot to work in specific channel this command is used in.')
async def setup(interaction: discord.Interaction):
    try:
        if active_sheets[str(interaction.guild.id)]:
            await interaction.response.send_message('setup has already been completed.')
    except:
        #copying sheet
        def copy_spreadsheet():
            copy_sheet = {
                'name': f'Pod Stats for: {interaction.guild.name}',
            }
            response = drive_service.files().copy(
                fileId= '1uHT4HWD_x00-AVKbeot7h-2OVcPnJfu-9y2cERcVmxU',
                body=copy_sheet
            ).execute()

            return response

        # Function to share the spreadsheet and generate a link
        def share_spreadsheet(file_id):
            # Set the file permissions to anyone with the link can view
            permission = {
                'type': 'anyone',
                'role': 'writer'
            }
            drive_service.permissions().create(
                fileId=file_id,
                body=permission
            ).execute()

            # Generate the link
            link = f"https://docs.google.com/spreadsheets/d/{file_id}/edit"
            return link

        channels[str(interaction.guild.id)] = interaction.channel.id
        save_channels(channels)
        await interaction.response.send_message(f'This channel has been set for the bot! use other slash (/) commands to use the bot\n'
                                                f'Making a stat tracking spreadsheet. . .')

        print('copying empty Pod Stat spreadsheet . . .')
        copiedFile = copy_spreadsheet()
        print('copied empty Pod Stat spreadsheet . . .')
        copiedFileID = copiedFile['id']

        active_sheets[str(interaction.guild.id)] = copiedFileID
        save_active_stats(active_sheets)

        link = share_spreadsheet(copiedFileID)
        print(f'sharing spreadsheet to {interaction.guild.name}')

        await interaction.followup.send(f'This channel has been set for the bot! use other slash (/) commands to use the bot\n'
                                                f"Here's the link: {link}\n"
                                                f'Open the link to start stat tracking!')

@tree.command(name= 'link', description= 'Sends a link for the google spreadsheet that is being used.')
async def send_link(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        await interaction.response.send_message(f'https://docs.google.com/spreadsheets/d/{active_sheets[str(interaction.guild.id)]}/?usp=sharing')
        print('shared link to google sheet')
    else:
        await interaction.response.send_message('This command must be run in the designated channel.')

@tree.command(name='addplayers', description= 'Adds players to the spreadsheet.')
async def addplayers(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        await interaction.response.send_message('Send the name of the player(s) you want to add to the spreadsheet\n'
                                                '(use commas in between names) or enter cancel to ignore command')

        def check(m):
            return m.author == interaction.user and m.channel == interaction.channel

        try:
            players = await bot.wait_for('message', check=check, timeout=90.0)
            players = players.content
            if players.lower() == 'cancel':
                await interaction.followup.send('Canceling . . .')
            else:
                player = players.split(', ') if ', ' in players else players.split(',')

                Title, Tables, RawData, Validations = get_sheets(active_sheets[str(interaction.guild.id)])
                spreadsheet = client.open(Title)
                Valid = spreadsheet.worksheet(title=Validations)
                playerCol = Valid.col_values(4)  # Get all values in the column


                # Update the column in Google Sheets
                update_range = f"D{len(playerCol) + 1}:D{len(playerCol) + len(player)}"
                update_data = [[play.capitalize()] for play in player]
                Valid.update(update_data, update_range)

                print(f'Added {len(player)} player(s) to spreadsheet {interaction.guild.name}')

                await interaction.followup.send(f'Added commander(s) {", ".join(player) if len(player) > 1 else "".join(player)} to spreadsheet')

        except asyncio.TimeoutError:
            await interaction.followup.send('You took too long to reply!')

@tree.command(name='addcommanders', description= 'Adds commanders to the spreadsheet.')
async def addcommanders(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        await interaction.response.send_message('Send the name of the commander(s) you want to add to the spreadsheet\n'
                                                '(use pipes ( | ) in between names) or enter cancel to ignore command')

        def check(m):
            return m.author == interaction.user and m.channel == interaction.channel

        try:
            commanders = await bot.wait_for('message', check=check, timeout=90.0)
            commanders = commanders.content
            if commanders.lower() == 'cancel':
                await interaction.followup.send('Canceling . . .')
            else:
                commander = commanders.split(' | ') if ' | ' in commanders else commanders.split('|')

                Title, Tables, RawData, Validations = get_sheets(active_sheets[str(interaction.guild.id)])
                spreadsheet = client.open(Title)
                Valid = spreadsheet.worksheet(title=Validations)
                commanderCol = Valid.col_values(1)  # Get all values in the column

                # Update the column in Google Sheets
                update_range = f"A{len(commanderCol) + 1}:A{len(commanderCol) + len(commander)}"
                update_data = [[play.capitalize()] for play in commander]
                Valid.update(update_data, update_range)

                print(f'Added {len(commander)} commander(s) to spreadsheet {interaction.guild.name}')

                await interaction.followup.send(f'Added commander(s) {", ".join(commander) if len(commander) > 1 else "".join(commander)} to spreadsheet')
        except asyncio.TimeoutError:
            await interaction.followup.send('You took too long to reply!')

@tree.command(name='addgame', description='Add a game to the spreadsheet, will ask for players and commanders.')
async def addGame(interaction: discord.Interaction):
    if is_active_game(interaction):
        await interaction.response.send_message('Use /finishgame to put in the results of previous game before starting a new game!')
    elif is_correct_channel(interaction) and not is_active_game(interaction):
        sheetID = active_sheets[str(interaction.guild.id)]

        Title, Tables, RawData, Validations = get_sheets(sheetID)
        spreadsheet = client.open(Title)
        allPlayers, allCommanders = data_refresh(spreadsheet.worksheet(title=Validations))

        if allPlayers == [] or allCommanders == []:
            await interaction.response.send_message('Please add some players and commanders first!')
            print('No players or commanders grabbed . . .')
        else:

            allPlayers    = '\n'.join(allPlayers)
            allCommanders = '\n'.join(allCommanders)

            await interaction.response.send_message(f'Please choose players:\n'
                                                    f'{allPlayers}\n'
                                                    f'Send a message with player names followed by a comma:')

            def check(m):
                return m.author == interaction.user and m.channel == interaction.channel

            try:
                players = await bot.wait_for('message', check=check, timeout=60.0)
                players = players.content

                if players.lower() == 'cancel':
                    await interaction.followup.send('Canceling . . .')
                else:
                    await interaction.followup.send(f'You selected: {players}\n'
                                                    f'\n'
                                                    f'Please choose commanders:\n'
                                                    f'{allCommanders}\n'
                                                    f'Send a message with commander names followed by a pipe ( | ):')

                    commanders = await bot.wait_for('message', check=check, timeout=60.0)
                    commanders = commanders.content
                    if commanders.lower() == 'cancel':
                        await interaction.followup.send('Canceling . . .')
                    else:
                        #putting players to commanders
                        player = players.split(', ') if ', ' in players else players.split(',')
                        commander = commanders.split(' | ') if ' | ' in commanders else commanders.split('|')

                        #after split, checks to see if theres a typo and corrects it
                        player = correct_name(player, allPlayers.split('\n'))
                        commander = correct_name(commander, allCommanders.split('\n'))

                        #joins player and commander together
                        playerAndCommander = [(f'{play} playing {comm}') for play, comm in zip(player, commander)]

                        #joining the players and commanders to make the message in discord look nice
                        gameInfo = '\n'.join(playerAndCommander)
                        await interaction.followup.send(f'You selected: {commanders}\n'
                                                        f'--------------------------\n'
                                                        f'{gameInfo}\n'
                                                        f'--------------------------\n'
                                                        f'Are both players and commanders correct?')
                        correct = await bot.wait_for('message', check=check, timeout=30.0)

                        if correct.content.lower().startswith('n'):
                            await interaction.followup.send('Ending transaction. Restart the command.')
                            print('Restarted input . . .')
                        elif correct.content.lower().startswith('y'):
                            await interaction.followup.send('Adding game to spreadsheet.')
                            print(f'Adding game in {interaction.guild.name} to spreadsheet {sheetID}. . .')

                            #sending data to google sheet
                            game_to_sheet(player, commander, spreadsheet.worksheet(title=RawData))
                            active_game[interaction.guild.id] = (player, commander)
                            print(active_game[interaction.guild.id])

            except asyncio.TimeoutError:
                await interaction.followup.send('You took too long to reply!')
    else:
        await interaction.response.send_message('This command must be run in the designated channel.')

@tree.command(name='finishgame', description='Finish a game by marking if a player died first, won the game and, got first blood (killed first).')
async def finishGame(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        if interaction.guild.id not in active_game:
            await interaction.response.send_message('No active game found.')

        players, commanders = active_game[interaction.guild.id]
        view = MyView(players, commanders)
        await interaction.response.send_message('Here are your buttons:', view=view)
    else:
        await interaction.response.send_message('This command must be run in the designated channel.')

@tree.command(name='tableplayer', description='Display the table for player stats.')
async def playerTable(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        #await interaction.response.send_message('Printing player table. . .')
        sheetID = active_sheets[str(interaction.guild.id)]

        Title, Tables, RawData, Validations = get_sheets(sheetID)
        spreadsheet = client.open(Title)

        tablesSheet = spreadsheet.worksheet(title=Tables)
        tableData = tablesSheet.get_all_values()

        playerNumber = tablesSheet.col_values(2)  # Get all values in the column
        lastPlayer = len(playerNumber)  # The last filled row is the length of the list

        players = pd.DataFrame(tableData)

        table1 = players.iloc[:lastPlayer]
        table1 = table1.iloc[1:, 1:8]
        table1_list = [table1.columns.tolist()] + table1.values.tolist()

        # Determine the maximum width for each column
        col_widths = [max(len(str(item)) for item in col) for col in zip(*table1_list)]

        # Create a formatted string for each row
        formatted_table = []

        # Send an initial response to acknowledge the interaction
        await interaction.response.defer()

        # Send a follow-up message to which you will make edits
        message = await interaction.followup.send("Starting to build the table...\n``` \n")

        # Assume table1_list is prepared, and col_widths calculated as before
        formatted_table = []

        for row in table1_list:
            formatted_row = " | ".join(f"{str(item):<{col_widths[i]}}" for i, item in enumerate(row))
            formatted_table.append(formatted_row)

            # Update the message with the current progress
            current_table = "\n".join(formatted_table)
            await message.edit(content=f"```\n{current_table}\n```")

            # Optional delay to simulate real-time updates
            await asyncio.sleep(0.5)  # Adjust the delay as needed

        # Finalize the message
        await message.edit(content=f"```\n{current_table}\n```")

@tree.command(name='tablecommander', description='Display the table for commander stats.')
async def commanderTable(interaction: discord.Interaction):
    if is_correct_channel(interaction):
        #await interaction.response.send_message('Printing player table. . .')
        sheetID = active_sheets[str(interaction.guild.id)]

        Title, Tables, RawData, Validations = get_sheets(sheetID)
        spreadsheet = client.open(Title)

        tablesSheet = spreadsheet.worksheet(title=Tables)
        tableData = tablesSheet.get_all_values()

        commanderNumber = tablesSheet.col_values(12)  # Get all values in the column
        lastPlayer = len(commanderNumber)  # The last filled row is the length of the list

        commanders = pd.DataFrame(tableData)

        table2 = commanders.iloc[:lastPlayer]
        table2 = table2.iloc[1:, 11:20]
        table2_list = [table2.columns.tolist()] + table2.values.tolist()

        # Determine the maximum width for each column
        col_widths = [max(len(str(item)) for item in col) for col in zip(*table2_list)]

        # Send an initial response to acknowledge the interaction
        await interaction.response.defer()

        # Send a follow-up message to which you will make edits
        message = await interaction.followup.send("Starting to build the table...\n``` \n")
        print(f'starting commanbder table for {sheetID} . . .')

        # Assume table1_list is prepared, and col_widths calculated as before
        formatted_table = []

        for row in table2_list:
            formatted_row = " | ".join(f"{str(item):<{col_widths[i]}}" for i, item in enumerate(row))
            formatted_table.append(formatted_row)

            # Update the message with the current progress
            current_table = "\n".join(formatted_table)
            await message.edit(content=f"```\n{current_table}\n```")

            # Optional delay to simulate real-time updates
            await asyncio.sleep(0.5)  # Adjust the delay as needed

        # Finalize the message
        await message.edit(content=f"```\n{current_table}\n```")
        print(f'finished commanbder table for {sheetID}')

@bot.event
async def on_ready():
    await tree.sync()
    print(f'We have logged in as {bot.user}')

if __name__ == '__main__':
    bot.run(discord_key)