from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import os
import os.path
import pickle
import base64
from email.mime.text import MIMEText
from email.header import decode_header
from datetime import datetime
import csv
import re
import html

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def get_message(service, user_id, msg_id):
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        msg_payload = message['payload']
        
        if 'parts' in msg_payload:
            parts = msg_payload['parts']
            email_body = ""
            for part in parts:
                if part['mimeType'] == 'text/plain':
                    data = part['body']['data']
                    decoded_data = base64.urlsafe_b64decode(data).decode('utf-8')
                    email_body += decoded_data
            return email_body
        else:
            return ""
    except Exception as error:
        print('An error occurred: %s' % error)
        return ""


def extract_email_address(email_body):
    # Regular expression pattern to match email addresses
    pattern = r'[\w\.-]+@[\w\.-]+'
    matches = re.findall(pattern, email_body)
    if matches:
        return matches[0]  # Return the first match
    else:
        return None





def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Build Gmail service
    service = build('gmail', 'v1', credentials=creds)

    # Check if CSV file exists and read existing email addresses
    existing_addresses = set()
    if os.path.exists('email_records.csv'):
        with open('email_records.csv', 'r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # Skip header row
            for row in reader:
                existing_addresses.add(row[1])  # Assuming email address is in the second column

    # Open CSV file in append mode
    with open('email_records.csv', 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        # Retrieve emails in batches
        page_token = None
        while True:
            results = service.users().messages().list(userId='me', pageToken=page_token).execute()
            messages = results.get('messages', [])

            if not messages:
                print('No more messages found.')
                break
            else:
                for message in messages:
                    msg_id = message['id']
                    email_body = get_message(service, 'me', msg_id)

                    # Extract email address from the email body
                    email_address = extract_email_address(email_body)

                    if email_address and email_address not in existing_addresses:
                        msg = service.users().messages().get(userId='me', id=msg_id).execute()  # Fetch the full message object
                        # Parse headers for date
                        headers = msg['payload']['headers']
                        date_str = next(header['value'] for header in headers if header['name'] == 'Date')
                        date = datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z')

                        # Format date and time in 12-hour format
                        datetime_12hr = date.strftime('%b %d, %Y, %I:%M %p')

                        # Write data to CSV file
                        writer.writerow([datetime_12hr, email_address])

                        # Add email address to set of existing addresses
                        existing_addresses.add(email_address)

            # Check if there are more pages of results
            page_token = results.get('nextPageToken')
            if not page_token:
                break

if __name__ == '__main__':
    main()