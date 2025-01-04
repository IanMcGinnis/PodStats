# Introduction

## What Is This Project?

This project uses Google Sheets API and a Discord bot to create Magic The Gathering stats for friends and a group of players (known as pods). I created this project to hopefully make it easier to insert data in a Google Spreedsheet for friends so I can get data for a statistics senior project in a year or two.

## How To Use It

To use the project, the code in `Main.py` needs to be ran locally after creating a [Google account for API](https://support.google.com/googleapi/answer/6158862?hl=en) and downloading a `.JSON` for permissions. After that go to Discords Development API and create a bot [here](https://www.writebots.com/discord-bot-token/) is how to create a bot. After getting a token, in the `example.env` copy it and make a `.env` to insert to the token.

# Current plans

## Phasing out google API

Google API wants the user to pay for the activity after about 600 API calls, so because of that I plan on changing the system from using google API to using a database and being able to output and provide data like I was for google API.

## MTG Card look up

I wanted to add looking up cards and providing pictures on this bot I realized it was too much to throw on this bot, I will make another bot that does card lookups and minigames like "what is this card?"
