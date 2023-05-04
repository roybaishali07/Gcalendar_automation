import csv
import datetime
import os.path
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build

# Define the color codes dictionary
color_codes = {
    'NLS': '5',  # BANANA
    'FLS' : '9',   #Blueberry
    'ES': '10',   # BASIL
    'EN': '6',   # TANGERINE
    'NS': '8',   # GRAPHITE
    'NN': '3',  # GRAPE
}

# Define file paths
CREDENTIALS_FILE = 'client_secret.json'
TOKEN_FILE = 'token.pickle'
INPUT_CSV_FILE = 'data/input.csv'
OUTPUT_FILE = 'data/output.txt'
OUTPUT_CSV_FILE = 'data/output.csv'

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


def getCSVData():
    # Read the input file in CSV format
    with open(INPUT_CSV_FILE, 'r') as f:
        reader = csv.DictReader(f)
        num_washes = 0
        events = []
        for row in reader:
            # Parse the start and end dates from the input row
            start_date_str = row['start_date']
            start_date_obj = datetime.datetime.strptime(start_date_str, '%Y-%m-%d').date()
            end_date_str = row['end_date/number_of_washes']
            if end_date_str.isdigit():
                num_washes = int(end_date_str)
            else:
                end_date_obj = datetime.datetime.strptime(end_date_str, '%Y-%m-%d').date()


            # Parse the start and end times from the input row
            start_time_str = row['start_time']
            start_time_obj = datetime.datetime.strptime(start_time_str, '%I:%M %p').time()
            end_time_str = row['end_time']
            end_time_obj = datetime.datetime.strptime(end_time_str, '%I:%M %p').time()

            # Parse the weekdays from the input row
            weekdays_str = row['weekdays']
            if '|' in weekdays_str:
                weekdays = weekdays_str.split('|')
            else:
                weekdays =  weekdays_str.split(',')

            # Parse the attendees and types from the input row
            if '|' in row['attendees']:
                attendees = row['attendees'].split('|')
            else:
                attendees = row['attendees'].split(',')

            if num_washes != 0:
                num_weeks = num_washes//len(weekdays)
                num_days = num_weeks * 7
                end_date_obj = start_date_obj + datetime.timedelta(days=num_days)
                num_washes = 0

            events.append[{
                "summary" : row['summary'], 
                "description" : row['description'], 
                "location" : row['location'], 
                "start_date" : start_date_obj, 
                "end_date" : end_date_obj, 
                "start_time" : start_time_obj, 
                "end_time" : end_time_obj, 
                "weekdays" : weekdays, 
                "attendees" : attendees, 
                "types" : row['type'], 
                "gmap_link" : row['gmap_link']}]

    return events        
    
def createEvents(service):
        events = get_events()

        for row in events:
            # Loop through each date in the range between start_date and end_date)
            current_date = row["start_date"]
            while current_date <= row["end_date"]:
                # Check if the current date's weekday matches the specified weekdays
                if current_date.strftime('%a') in row["weekdays"]:
                    # Create a new event on Google Calendar
                    start_datetime = datetime.datetime.combine(current_date, row["start_time"])
                    end_datetime = datetime.datetime.combine(current_date, row["end_time"])
                    event = {
                        'summary': row['summary'],
                        'description': f"{row['description']} - View location on Google Maps: {row['gmap_link']}",
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
                        'colorId': color_codes.get(row["type"], None)
                    }
                    #Add each attendee to the event
                    for i, attendee in enumerate(row["attendees"]):
                        if attendee and '@' in attendee:
                            attendee_obj = {
                                'email': attendee,
                                'displayName': attendee.split('@')[0],
                                'responseStatus': 'needsAction'
                            }
                            event['attendees'].append(attendee_obj)
                    event = service.events().insert(calendarId='primary', body=event).execute()

                    with open(OUTPUT_FILE, 'a') as f:
                        f.write(f'Event created for {current_date_obj}: {event["htmlLink"]}\n')
                # Move to the next date
                current_date_obj += datetime.timedelta(days=1)

def get_events(start_date, end_date):
    creds = authenticate()
    service = build(API_SERVICE_NAME, API_VERSION, credentials=creds)
    events_result = None
    try:
        events_result = service.events().list(
            calendarId='primary',
            timeMin=start_date.isoformat() + 'Z',
            timeMax=end_date.isoformat() + 'Z',
            maxResults=1000,
            singleEvents=True,
            orderBy='startTime',
            fields='items(summary,description,location,start,end,attendees,colorId)'
        ).execute()
    except HttpError as error:
        print(f'An error occurred: {error}')
    events = events_result.get('items', [])
    return events

def read_dates():
    start_date_str = input("Enter start date in YYYY-MM-DD format: ")
    end_date_str = input("Enter end date in YYYY-MM-DD format: ")
    start_date = datetime.datetime.strptime(start_date_str, '%Y-%m-%d')
    end_date = datetime.datetime.strptime(end_date_str, '%Y-%m-%d')
    return start_date, end_date

def write_events(events):
    with open(OUTPUT_CSV_FILE, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['date', 'start_time', 'end_time', 'weekday', 'summary', 'description', 'attendees', 'location'])
        for event in events:
            start_time = datetime.datetime.fromisoformat(event['start'].get('dateTime', event['start'].get('date')))
            end_time = datetime.datetime.fromisoformat(event['end'].get('dateTime', event['end'].get('date')))
            date_str = start_time.date().isoformat()
            start_time_str = start_time.time().strftime('%H:%M')
            end_time_str = end_time.time().strftime('%H:%M')
            weekday_str = start_time.strftime('%A')
            summary = event['summary']
            description = event.get('description', 'No description provided')
            location = event.get('location', 'No location provided')
            attendees = ', '.join([attendee['email'] for attendee in event.get('attendees', [])])
            writer.writerow([date_str, start_time_str, end_time_str, weekday_str, summary, description, attendees, location])





def main():
    try:
        creds = authenticate()
        service = build(API_SERVICE_NAME, API_VERSION, credentials=creds)
        
    except HttpError as e:
        print(e)
        print('Authentication failed. Please try again.')

    while True:
        choice = input("""
        \n1. Create a new events from csv file.
2. Fetch all events from calendar.
0.Exit.
        
Enter your choice : """)

        if choice == '1':
            createEvents(service)
            continue
        
        elif choice == '2':
            start_date, end_date = read_dates()
            events = get_events(start_date, end_date)
            write_events(events)
            print(f'Events written to {OUTPUT_CSV_FILE}')
            continue
        
        elif choice == '0':
            break

main()
        
