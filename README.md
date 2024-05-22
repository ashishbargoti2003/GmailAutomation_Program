
# Gmail Automation Script

## Overview
This script automates the process of sending emails using Gmail and integrates with Google Sheets to fetch recipient details and update the status of sent emails.

## About the Project
The Gmail Automation Script is designed to streamline email communication for organizations or individuals who need to send out multiple personalized emails efficiently. This is particularly useful for marketing campaigns, customer engagement, or any situation requiring the dispatch of bulk emails with customized content.

## Key Features
- **Email Automation:** Automates the sending of emails, reducing manual effort and minimizing the risk of errors.
- **Google Sheets Integration:** Reads recipient details such as names and email addresses from a Google Sheet, allowing easy management and updates of recipient lists.
- **Personalized Content:** Customizes email content by inserting recipient-specific information, ensuring a personal touch in each email.
- **Status Updates:** Updates the Google Sheet with the timestamp of when each email was sent, providing a clear record of communication.
- **Secure Connection:** Utilizes SSL to ensure that all email communications are sent securely.
- **Error Handling:** Includes basic error handling to manage issues such as failed email sends, ensuring robustness and reliability.
- **Customizable Template:** Allows for the customization of email content via an external text file, making it easy to modify messages without changing the code.
- **Efficient Workflow:** Fetches recipient details and sends emails in a loop, efficiently processing multiple records in one execution.

## Configuration

## Dependencies
The script relies on several Python libraries:
- `email.message.EmailMessage` for constructing email messages
- `ssl` for secure connections
- `smtplib` for sending emails via SMTP
- `gspread` for interacting with Google Sheets
- `oauth2client.service_account.ServiceAccountCredentials` for Google API authentication
- `datetime` for handling date and time

## Scope
This script performs the following tasks:
1. Authorizes access to Google Sheets using service account credentials.
2. Reads recipient details from a Google Sheet.
3. Sends customized emails to each recipient.
4. Updates the Google Sheet with the timestamp of when the email was sent.

## Configuration

### Google API Setup
1. Ensure you have a Google Cloud project with the Google Sheets API and Google Drive API enabled.
2. Obtain a service account key JSON file and place it in the specified directory (`Project/Gmail_Automation/Data/key.json`).

### Google Sheet Setup
1. Create a Google Sheet named "GmailAutomation".
2. Ensure the first sheet (Sheet1) has columns: `name`, `email`, `status`.
3. Populate the sheet with recipient names and emails.

## Script Breakdown

### Import Libraries
```python
from email.message import EmailMessage
import ssl
import smtplib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
```
## How to run:
- Clone GitHub repository
- Open the folder inside VS Code
- Activate virtual environment by migrating inside scripts and activating "activate.ps1"
- Click on run button

