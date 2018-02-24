"""
~ NOTES ~

directories.py is for email and Google Calendar DICTIONARIES.
utils.py is for FUNCTIONS
variables.py is for OTHER VARIABLES (mostly for Google Sheets)

1. Google Sheets Interactions
    This file contains the hard-coded Google Sheets cell ranges that the
    bot should be looking at. Further explanation is located above the
    variables, and description of the Google Sheets setup is located
    elsewhere.
"""

import os

# This is for saving cached schedules and whatnot
LOCAL_PATH = "/Users/keatonarmentrout/Desktop/Techstars/Mentor Madness/"


# Login credentials for MM Bot
mm_bot_gmail_name = 'mentor.madness.bot@gmail.com'
mm_bot_gmail_password = 'aZ7Z^HPe##XPiz07'


# Ideally the login and password should be in environment variables
# TODO: enable this ^
# gmail_credentials = {
#     'name': os.environ['mm_bot_gmail_name'],
#     'password': os.environ['mm_bot_gmail_password']
# }


'''
This is the number of days the update script should be looking ahead of
the current day. E.g. if lookahead_days == 2, then people will get alerts
when if anything is changed on either the current day or the 2 days
following (of those days in `sheet_options` below. Uses that index, not
calendar days). This should be no more than 2, just to limit spam.
'''
lookahead_days = 1


# Google Sheets spreadsheet ID
spreadsheet_id = '18gb1ehs9-hmXbIkKaTcLUvurzAJzpjDiXgNFZeazrNA'  # This is the MM spreadsheet
test_spreadsheet_id = '1qdAgkuyAl6DRV3LRn-zheWSiD-r4JIya8Ssr6-DswY4'  # This is my test spreadsheet


'''
This is a list of the spreadsheet tab names. For clarity, they should
all follow the same naming convention and contain the date.
'''
sheet_options = [
    'Mon 2/13', 'Tues 2/14', 'Wed 2/15', 'Thurs 2/16', 'Fri 2/17',
    'Mon 2/20', 'Tues 2/21', 'Wed 2/22', 'Thurs 2/23', 'Fri 2/24',
    'Mon 2/27', 'Tues 2/28', 'Wed 3/1', 'Thurs 3/2', 'Fri 3/3',
    'Mon 3/6', 'Tues 3/7', 'Wed 3/8', 'Thurs 3/9', 'Fri 3/10'
]


# Break up the days into weeks
week1 = ['Mon 2/13', 'Tues 2/14', 'Wed 2/15', 'Thurs 2/16', 'Fri 2/17']
week2 = ['Mon 2/20', 'Tues 2/21', 'Wed 2/22', 'Thurs 2/23', 'Fri 2/24']
week3 = ['Mon 2/27', 'Tues 2/28', 'Wed 3/1', 'Thurs 3/2', 'Fri 3/3']
week4 = ['Mon 3/6', 'Tues 3/7', 'Wed 3/8', 'Thurs 3/9', 'Fri 3/10']


'''
Room mapping should consist of a dict of dicts,
where each subdict has "name", "mentor_col", and
"check_range" fields.
"name" == name of the room (if applicable)
"mentor_col" == column in Google Sheets containing the name of
    the mentor assigned to that room
"check_range" == columns in Google Sheets containing the names
    of companies and associates in that room

Note that this bot doesn't check for changes in the mentor names.
'''
room_mapping = {
        1: {'name': 'Glacier', 'mentor_col': 1, 'check_range': [2, 3]},
        2: {'name': 'Harbor', 'mentor_col': 4, 'check_range': [5, 6]},
        3: {'name': 'Lagoon', 'mentor_col': 7, 'check_range': [8, 9]},
        4: {'name': 'Abyss', 'mentor_col': 10, 'check_range': [11, 12]},
        5: {'name': 'Classroom', 'mentor_col': 13, 'check_range': [14, 15]},
        6: {'name': 'Across Hall', 'mentor_col': 16, 'check_range': [17, 18]}
    }


mentor_columns = [1, 4, 7, 10, 13, 16]

start_col = 'A'
end_col = 'S'




'''
This variable explicitly sets the length of each Google Sheets row.
The formula it uses is based on assuming 3 columns for each room
listed in `room_mapping` (to list mentor, company, and associate),
plus an additional column on the left to hold the time slot.

CHANGE THIS FORMULA if the spreadsheet is set up differently.
'''
row_length = len(room_mapping)*3 + 1


'''
This should include every cell on the Google Sheets page
that the bot should look at. Includes the mentor, company,
and associate columns for each of the rooms, as well as time slots.
'''
full_range = 'A1:S20'


headers = [
    'room1_mentor', 'room1_company', 'room1_associate',
    'room2_mentor', 'room2_company', 'room2_associate',
    'room3_mentor', 'room3_company', 'room3_associate',
    'room4_mentor', 'room4_company', 'room4_associate',
    'room5_mentor', 'room5_company', 'room5_associate',
    'room6_mentor', 'room6_company', 'room6_associate',
]

value_input_option = 'RAW'

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
