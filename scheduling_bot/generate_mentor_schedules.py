import os

import email_sender
import variables as vrs
import utils


def main(specific_day=None):
    # Make Google API object
    sheets_api = utils.google_sheets_login()

    # Set variables
    spreadsheet_id = vrs.spreadsheet_id
    room_mapping = vrs.room_mapping
    full_range = vrs.full_range

    sheet_names = []
    if not specific_day:
        sheet_names = [utils.get_next_day()]
    else:
        if isinstance(specific_day, list):
            sheet_names = specific_day
        elif isinstance(specific_day, str):
            sheet_names = [specific_day]

    mentor_dict = {}

    for day in sheet_names:
        # String formatting for API query and file saving
        sheet_query = day + '!' + full_range

        # Make request for sheet
        sheet = sheets_api.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=sheet_query).execute()
        new_sheet = sheet['values']
        new_sheet = new_sheet[1:]  # Get rid of the header row

        # Make sure each row is the correct length
        for idx, new_sheet_row in enumerate(new_sheet):
            if len(new_sheet_row) < vrs.row_length:
                new_sheet[idx].extend([''] * (vrs.row_length - len(new_sheet_row)))

        for row in new_sheet:
            timeslot = row[0]

            # Iterate over rooms
            for room_num in range(1, len(room_mapping) + 1):
                # Get descriptive variables of room
                room_dict = room_mapping[room_num]
                room_name = room_dict['name']
                mentor_name = row[room_dict['mentor_col']]

                if not mentor_name:
                    continue

                # Add mentor to mentor list for the day
                if mentor_dict.get(mentor_name) is None:
                    mentor_dict[mentor_name] = []

                teamname_idx = room_dict['check_range'][0]
                teamname = utils.process_name(row[teamname_idx])
                if teamname:
                    new_event_dict = {'time': timeslot, 'mentor': mentor_name, 'company': teamname,
                                      'room_num': str(room_num), 'room_name': room_name, 'day': day}
                    mentor_dict[mentor_name].append(new_event_dict)

    email_sender.make_daily_mentor_schedules(mentor_dict)
    email_sender.make_mentor_packet_schedules(mentor_dict)


if __name__ == '__main__':
    main()
