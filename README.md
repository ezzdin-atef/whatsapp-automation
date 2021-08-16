# WhatsApp Automation

WhatsApp Automation Script for Sending Messages

## Prerequirement

You need to run this script `Node.js` and you can download it from [here](https://nodejs.org)

## NPM Packages we used

- `puppeteer` for browser automation
- `async` for better way for woking with asynchronous JavaScript.
- `read-excel-file` for reading excel files

## How to use

1. Download the file as `ZIP` or with git command `git clone`
1. Open the terminal in the project directory and type `npm install` or `npm i`
1. Edit the contact excel file with the new numbers
1. Generate the VCard file and import it in your phone
1. Change the message in `index.js` file
1. Run the script by typing `npm start`
1. Finally, Scan the QR Code to connect the browser with Web WhatsApp

## How to generate VCard file

1. Run the script by typing `npm run vcardgen`
1. You will find `vcard.vcf` in the root directory, so import it in your phone

## How to generate VCard file from excel

1. Open the developer tab in excel
1. Open Visual Basic and paste the code from `THE VBA SCRIPT` file
1. Run it

Hope you like it. If you do, Please star ⭐ the repo ❤
