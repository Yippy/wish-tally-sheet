# Wish Tally
Script to help manage Genshin Impact wish history using Google Sheet Document.

## Project Website
Visit the Genshin Impact collection of Google Sheets tools:

https://gensheets.co.uk 

## Google Add-on
Wish Tally is available on Google Workspace Marketplace from 6th November 2021.
https://workspace.google.com/marketplace/app/wish_tally/791037722195

## Preview
<img src="https://raw.github.com/Yippy/wish-tally-sheet/master/images/wish-tally-preview.png?sanitize=true">

## Tutorial
[Install Add-on](docs/INSTALL_ADD_ON.md)


## Template Document
If you prefer to use the Wish Tally document with embedded script, you can make a copy here:
https://docs.google.com/spreadsheets/d/1_Or0KRVZ5nwCrHdO5c_8rqu2CWJ_aLETnZBLSYBDS_c/edit

## How to compile script
This project uses https://github.com/google/clasp to help compile code to Google Script.

1. Run ```npm install -g @google/clasp```
2. Edit the file .clasp.json with your Google Script
3. Run ```clasp login```
4. Run ```clasp push -w```