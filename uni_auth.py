from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Encoding
from base64 import urlsafe_b64encode

# Email libraries
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://mail.google.com/']

def build_message(destination, sender, subject, name, your_name, customized, body):
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = destination
    msg['Subject'] = subject
    text = body.format(name, customized, your_name)
    msg.attach(MIMEText(text, 'plain'))
    return {'raw': urlsafe_b64encode(msg.as_bytes()).decode()}

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)

        # Have user input all relevant details
        sender = str(input('Your Email:'))
        Subject = str(input('Subject line:'))
        your_name = input("What is your name:")

        # Load in text file with all the messages
        text_file = open("message.txt", "r")
        text_read = text_file.read()
        
        # Parse the spreadsheet
        wb = xl.load_workbook(r'email_list.xlsx')
        sheet1 = wb['Sheet1']

        # Array Definitions
        names = []
        emails = []
        subjects = []
        customization = []

        # Fill the arrays
        for cell in sheet1['A']:
            names.append(cell.value)
        names = list(filter(None, names))

        for cell in sheet1['B']:
            emails.append(cell.value)
        emails = list(filter(None, emails))

        for cell in sheet1['C']:
            subjects.append(cell.value)
        subjects = list(filter(None, subjects))

        for cell in sheet1['D']:
            customization.append(cell.value)
        customization = list(filter(None, customization))

        # Send the emails
        for i in range(len(names)):
            full_msg = build_message(destination = emails[i], 
                                     sender = sender, 
                                     subject = Subject + subjects[i], 
                                     name = names[i], 
                                     your_name = your_name, 
                                     customized = customization[i], 
                                     body = text_read)

            try:
                service.users().messages().send(userId = "me",
                                                body = full_msg).execute()
                print('Mail sent to', emails[i])
            except:
                print("Unable to send email to", emails[i])
            
        print(len(names), 'Emails sent successfully')

    except HttpError as error:
        print("An error occurred:", error)


if __name__ == '__main__':
    main()