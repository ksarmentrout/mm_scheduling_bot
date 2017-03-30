import re
import os
from datetime import datetime, timedelta

from httplib2 import Http
from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials


def parse_webhook_json(added_json):
    ad = added_json
    ad['name'] = ad.get('first_name') + ' ' + ad.get('last_name')
    time_dict = parse_time(ad)
    ad.update(**time_dict)

    return ad


def parse_time(original_dict):
    time_dict = {}
    start_field = original_dict['start_time']
    end_field = original_dict['end_time']

    # Get the day
    day = re.search('([0-9]{1,2}/[0-9]{1,2})/', start_field)
    if not day:
        return {}

    time_dict['day'] = day.group(1)

    # Get the start and end times
    start_time = re.search('\s([0-9]{1,2}:[0-9]{2})', start_field)
    time_dict['start_time'] = start_time.group(1)

    end_time = re.search('\s([0-9]{1,2}:[0-9]{2})', end_field)
    time_dict['end_time'] = end_time.group(1)

    return time_dict


def google_sheets_login():
    # Get credentials from Google Developer Console
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    # json_path = "/Users/keatonarmentrout/Desktop/Techstars/Mentor Madness/scheduling_bot/MM-Scheduler-61cb94ec8350.json"
    credentials = ServiceAccountCredentials.from_json_keyfile_name('functions/MM_Bot-32fa78cfd51b.json', scopes=scopes)

    # Authenticate using Http object
    http_auth = credentials.authorize(Http())

    # Build Google API response object for sheets
    sheets_api = build('sheets', 'v4', credentials=credentials)

    return sheets_api


def google_calendar_login():
    # Get credentials from Google Developer Console
    scopes = ['https://www.googleapis.com/auth/calendar']

    # NOTE: The 'MM-Scheduler' service account is owned by the mentor.madness.bot. The 'MM-Bot' service account
    # is owned by my personal email.
    json_path = "/Users/keatonarmentrout/Desktop/Techstars/Mentor Madness/scheduling_bot/MM-Scheduler-61cb94ec8350.json"
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_path, scopes=scopes)

    # Authenticate using Http object
    http_auth = credentials.authorize(Http())

    # Build Google API response object for sheets
    sheets_api = build('calendar', 'v3', credentials=credentials)

    return sheets_api


def make_cell_range(start_time, end_time):
    start_bound = spreadsheet_time_mapping.get(start_time)
    end_bound = spreadsheet_time_mapping.get(end_time)
    end_bound = str(int(end_bound) - 1)
    cell_range = start_col + start_bound + ':' + end_col + end_bound
    return cell_range


gmail_credentials = {
    'name': os.environ['mm_bot_gmail_name'],
    'password': os.environ['mm_bot_gmail_password']
}

spreadsheet_id = '18gb1ehs9-hmXbIkKaTcLUvurzAJzpjDiXgNFZeazrNA'

room_mapping = {
        1: {'name': 'Glacier', 'mentor_col': 1, 'check_range': [2, 3]},
        2: {'name': 'Harbor', 'mentor_col': 4, 'check_range': [5, 6]},
        3: {'name': 'Lagoon', 'mentor_col': 7, 'check_range': [8, 9]},
        4: {'name': 'Abyss', 'mentor_col': 10, 'check_range': [11, 12]},
        5: {'name': 'Classroom', 'mentor_col': 13, 'check_range': [14, 15]},
        6: {'name': 'Across Hall', 'mentor_col': 16, 'check_range': [17, 18]}
    }

mentor_columns = [1, 4, 7, 10, 13, 16]

full_range = 'A1:S20'

sheet_options = [
        'Mon 2/13', 'Tues 2/14', 'Wed 2/15', 'Thurs 2/16', 'Fri 2/17',
        'Mon 2/20', 'Tues 2/21', 'Wed 2/22', 'Thurs 2/23', 'Fri 2/24',
        'Mon 2/27', 'Tues 2/28', 'Wed 3/1', 'Thurs 3/2', 'Fri 3/3'
    ]

spreadsheet_time_mapping = {
    '9:00': '2',
    '9:30': '3',
    '10:00': '4',
    '10:30': '5',
    '11:00': '6',
    '11:30': '7',
    '12:00': '8',
    '12:30': '9',
    '1:00': '10',
    '1:30': '11',
    '2:00': '12',
    '2:30': '13',
    '3:00': '14',
    '3:30': '15',
    '4:00': '16',
    '4:30': '17',
    '5:00': '18',
    '5:30': '19',
    '6:00': '20'
}

start_col = 'A'
end_col = 'S'

headers = [
    'room1_mentor', 'room1_company', 'room1_associate',
    'room2_mentor', 'room2_company', 'room2_associate',
    'room3_mentor', 'room3_company', 'room3_associate',
    'room4_mentor', 'room4_company', 'room4_associate',
    'room5_mentor', 'room5_company', 'room5_associate',
    'room6_mentor', 'room6_company', 'room6_associate',
]

# This is based on the number of rooms. If there are 6 rooms,
# each room is made up of 3 columns. Then including the time
# column on the left makes row_length = 3 * (number of rooms) + 1
row_length = 19

value_input_option = 'RAW'


def make_time_range(start_time):
    # Creates a half-hour time window based on the start time

    # Pads with zero
    if len(start_time[:start_time.find(' ')]) == 4:
        start_time = '0' + start_time

    # Shifts to 24-hour clock
    if 'PM' in start_time and start_time[:2] != '12':
        hour = int(start_time[:2])
        hour = str(hour + 12)
        start_time = hour + start_time[2:]

    start_time = start_time[:start_time.find(' ')]

    start_obj = datetime.strptime(start_time, '%H:%M')
    end_obj = start_obj + timedelta(minutes=30)
    end_time = end_obj.strftime('%H:%M')

    time_range = start_time + '-' + end_time
    return time_range


def process_name(original_name):
    # ALL CASES need to be checked because sometimes other information
    # besides just the name goes into these boxes. I know it's ugly.
    name = original_name.lower()
    if 'rate' in name:
        name = 'rate'
    elif 'alice' in name:
        name = 'alice'
    elif 'care' in name:
        name = 'care'
    elif 'machine' in name:
        name = 'sea machines'
    elif name.strip() == 'sea':
        name = 'sea machines'
    elif 'brain' in name:
        name = 'brainspec'
    elif 'brizi' in name:
        name = 'brizi'
    elif 'evolve' in name:
        name = 'evolve'
    elif 'lorem' in name:
        name = 'lorem'
    elif 'nix' in name:
        name = 'nix'
    elif 'offgrid' in name:
        name = 'offgridbox'
    elif 'solstice' in name:
        name = 'solstice'
    elif 'tive' in name:
        name = 'tive'
    elif 'voatz' in name:
        name = 'voatz'
    elif 'ashley' in name:
        name = 'ashley'
    elif 'andrew' in name:
        name = 'andrew'
    elif 'dmitriy' in name:
        name = 'dmitriy'
    elif 'dimitry' in name:
        name = 'dmitriy'
    elif 'justin' in name:
        name = 'justin'
    elif 'keaton' in name:
        name = 'keaton'
    elif 'lillian' in name:
        name = 'lillian'
    elif 'lilian' in name:
        name = 'lillian'
    elif 'troy' in name:
        name = 'troy'
    elif 'yada' in name:
        name = 'yada'
    elif 'max' in name:
        name = 'max'
    elif 'dan' in name:
        name = 'dan'
    elif 'miranda' in name:
        name = 'miranda'
    else:
        name = 'not_found'
    return name

