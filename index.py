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

# Define the color codes dictionary
color_codes = {
    'NLS': '5',  # BANANA
    'FLS' : '9',   #Blueberry
    'ES': '10',   # BASIL
    'EN': '6',   # TANGERINE
    'NS': '8',   # GRAPHITE
    'NN': '3',  # GRAPE
}

# Read the input file in CSV format
with open('input.csv', 'r') as f:
    reader = csv.DictReader(f)
    for row in reader:
        # Parse the start and end dates from the input row
        start_date_str = row['start_date']
        start_date_obj = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date_str = row['end_date']
        end_date_obj = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date()

        # Parse the start and end times from the input row
        start_time_str = row['start_time']
        start_time_obj = datetime.datetime.strptime(start_time_str, '%I%p').time()
        end_time_str = row['end_time']
        end_time_obj = datetime.datetime.strptime(end_time_str, '%I%p').time()

        # Parse the weekdays from the input row
        weekdays_str = row['weekdays']
        weekdays = [x.strip()[:3] for x in weekdays_str.split('|')]

        # Parse the attendees and types from the input row
        attendees = row['attendees'].split('|')
        types = row['type'].split('|')

        # Loop through each date in the range between start_date and end_date
        current_date_obj = start_date_obj
        while current_date_obj <= end_date_obj:
            # Check if the current date's weekday matches the specified weekdays
            if current_date_obj.strftime('%a') in weekdays:
                # Create a new event on Google Calendar
                start_datetime = datetime.datetime.combine(current_date_obj, start_time_obj)
                end_datetime = datetime.datetime.combine(current_date_obj, end_time_obj)
                event = {
                    'summary': row['summary'],
                    'description': row['description'],
                    'start': {
                        'dateTime': start_datetime.isoformat(),
                        'timeZone': 'Asia/Kolkata',
                    },
                    'end': {
                        'dateTime': end_datetime.isoformat(),
                        'timeZone': 'Asia/Kolkata',
                    },
                    'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'email', 'minutes': 60},
                            {'method': 'popup', 'minutes': 30},
                        ],
                    },
                    'attendees': [],
                    'location': row['location'],
                    'colorId': color_codes.get(types[0], None)
                }
                # Add each attendee to the event
                for i, attendee in enumerate(attendees):
                    if attendee and '@' in attendee:
                        attendee_obj = {
                            'email': attendee,
                            'displayName': attendee.split('@')[0],
                            'responseStatus': 'needsAction'
                        }
                        event['attendees'].append(attendee_obj)
                event = service.events().insert(calendarId='primary', body=event).execute()

                with open('output.txt', 'a') as f:
                    f.write(f'Event created for {current_date_obj}: {event["htmlLink"]}\n')

            #Move to the next date
            current_date_obj += datetime.timedelta(days=1)


