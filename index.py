import csv
import datetime
import os.path
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build

# Define file paths
CREDENTIALS_FILE = 'client_secret.json'
TOKEN_FILE = 'token.pickle'
INPUT_FILE = 'input.csv'
OUTPUT_FILE = 'output.txt'

# Define API scopes and version
API_SERVICE_NAME = 'calendar'
API_VERSION = 'v3'
API_SCOPES = ['https://www.googleapis.com/auth/calendar.readonly', 'https://www.googleapis.com/auth/calendar.events']

def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, API_SCOPES)
            creds = flow.run_local_server(port=41519)
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return creds

creds = authenticate()
service = build(API_SERVICE_NAME, API_VERSION, credentials=creds)

# Read the input file in CSV format
with open('input.txt', 'r') as f:
    reader = csv.DictReader(f)
    for row in reader:
        # Parse the date from the input row
        date_str = row['date']
        date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()

        # Set the start time of the event to 10 am on the specified date
        start_time = datetime.datetime.combine(date_obj, datetime.time(10, 0, 0))

        # Split the emails field by the | character and create a list of attendees
        emails = row['emails'].split('|')
        attendees = [{'email': email} for email in emails if email and '@' in email]

        # Set the event location
        location = row['location']

        # Create a new event on Google Calendar
        event = {
            'summary': row['summary'],
            'description': row['description'],
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': 'Asia/Kolkata',
            },
            'end': {
                'dateTime': (start_time + datetime.timedelta(hours=1)).isoformat(),
                'timeZone': 'Asia/Kolkata',
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 60},
                    {'method': 'popup', 'minutes': 30},
                ],
            },
            'attendees': attendees,
            'location': location,
        }
        event = service.events().insert(calendarId='primary', body=event).execute()

        # Write the event details to the output file
        with open('output.txt', 'a') as f:
            f.write(f'Event created: {event["htmlLink"]}\n')
