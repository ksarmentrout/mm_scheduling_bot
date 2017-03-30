import csv
import smtplib
from email.mime.text import MIMEText

import directories as dr
import daily_notice
import utils


def send_added_msgs(msg_dicts, server):
    for event in msg_dicts:
        clean_name = utils.process_name(event['name'])

        # Get full schedule for the day for that team or associate
        full_schedule = daily_notice.main(team=clean_name, specific_day=event['day'], send_emails=False)
        if clean_name == 'not_found':
            address_name = event['name']
        else:
            address_name = dr.names[clean_name]
        msg = 'Hello ' + address_name + ',\n\n' + \
              'You have been scheduled for the following meeting with ' + event['mentor'] + ':\n\n' + \
              event['day'] + ' - ' + event['time'] + '\n' + \
              'Room ' + event['room_num'] + ' (' + event['room_name'] + ')\n\n\n\n' + \
              'Your full updated schedule for ' + event['day'] + ' is as follows:'

        event_list = bulk_event_formatter(full_schedule[clean_name])
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += ''.join(event_list)

        msg += '\n\nPlease check the main schedule if this is in error. Have a great meeting!\n\n' + \
        '- Scheduling Bot'

        to_addresses = dr.update_email_list[clean_name]

        if isinstance(to_addresses, list):
            for addr in to_addresses:
                message = MIMEText(msg)
                message['From'] = 'mentor.madness.bot@gmail.com'
                message['To'] = addr
                message['Subject'] = 'New mentor meeting at ' + event['time'] + ' on '+ event['day']

                server.send_message(message)
                print('Sent added email - ' + address_name + ' with ' + event['mentor'] + ' at ' + event['time'])
        else:
            message = MIMEText(msg)
            message['From'] = 'mentor.madness.bot@gmail.com'
            message['To'] = to_addresses
            message['Subject'] = 'New mentor meeting at ' + event['time'] + ' on '+ event['day']

            server.send_message(message)
            print('Sent added email - ' + address_name + ' with ' + event['mentor'] + ' at ' + event['time'])


def send_deleted_msgs(msg_dicts, server):
    for event in msg_dicts:
        clean_name = utils.process_name(event['name'])

        # Get full schedule for the day for that team or associate
        full_schedule = daily_notice.main(team=clean_name, specific_day=event['day'], send_emails=False)

        if clean_name == 'not_found':
            address_name = event['name']
        else:
            address_name = dr.names[clean_name]
        msg = 'Hello ' + address_name + ',\n\n' + \
              'You have been REMOVED FROM the following meeting with ' + event['mentor'] + ':\n\n' + \
              event['day'] + ' - ' + event['time'] + '\n' + \
              'Room ' + event['room_num'] + ' (' + event['room_name'] + ')\n\n\n\n' + \
              'Your full updated schedule for ' + event['day'] + ' is as follows:'

        event_list = bulk_event_formatter(full_schedule[clean_name])
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += ''.join(event_list)

        msg += '\n\nPlease check the main schedule if this is in error.\n\n' + \
               '- Scheduling Bot'

        to_addresses = dr.update_email_list[clean_name]

        if isinstance(to_addresses, list):
            for addr in to_addresses:
                message = MIMEText(msg)
                message['From'] = 'mentor.madness.bot@gmail.com'
                message['To'] = addr
                message['Subject'] = 'Cancelled mentor meeting at ' + event['time'] + ' on ' + event['day']

                server.send_message(message)
                print('Sent deleted email - ' + address_name + ' with ' + event['mentor'] + ' at ' + event['time'])
        else:
            message = MIMEText(msg)
            message['From'] = 'mentor.madness.bot@gmail.com'
            message['To'] = to_addresses
            message['Subject'] = 'Cancelled mentor meeting at ' + event['time'] + ' on ' + event['day']

            server.send_message(message)
            print('Sent deleted email - ' + address_name + ' with ' + event['mentor'] + ' at ' + event['time'])


def send_update_mail(added_msg_dicts, deleted_msg_dicts):
    server = email_login()

    # Send mail
    send_added_msgs(added_msg_dicts, server)
    send_deleted_msgs(deleted_msg_dicts, server)

    # Close connection
    server.quit()


def send_daily_mail(targets):
    server = email_login()

    for key, events in targets.items():
        msg = 'Hello ' + dr.names[key] + ',\n\n' + \
              'Here are your scheduled meetings for tomorrow:\n\n'

        event_list = bulk_event_formatter(events)
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += ''.join(event_list)

        msg += '\n\nPlease contact Ty or check the main schedule if this is in error.\n\n' + \
              '- Scheduling Bot'

        to_addresses = dr.daily_email_list[key]

        if isinstance(to_addresses, list):
            for addr in to_addresses:
                message = MIMEText(msg)
                message['From'] = 'mentor.madness.bot@gmail.com'
                message['To'] = addr
                message['Subject'] = 'Mentor meeting summary for tomorrow'

                server.send_message(message)
                print('Sent daily email to ' + key)
        else:
            message = MIMEText(msg)
            message['From'] = 'mentor.madness.bot@gmail.com'
            message['To'] = to_addresses
            message['Subject'] = 'Mentor meeting summary for tomorrow'

            server.send_message(message)
            print('Sent daily email to ' + key)

    server.quit()


def send_weekly_mail(targets):
    server = email_login()

    for key, events in targets.items():
        msg = 'Hello ' + dr.names[key] + ',\n\n' + \
              'Here are your scheduled meetings for this week:\n\n'

        event_list = bulk_event_formatter(events)
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += ''.join(event_list)

        msg += '\n\nThis represents the first draft for the week. Please check the main schedule if this is in error.\n\n' + \
              '- Scheduling Bot'

        to_addresses = dr.daily_email_list[key]

        if isinstance(to_addresses, list):
            for addr in to_addresses:
                message = MIMEText(msg)
                message['From'] = 'mentor.madness.bot@gmail.com'
                message['To'] = addr
                message['Subject'] = 'Mentor meeting summary for next week'

                server.send_message(message)
                print('Sent weekly email to ' + key)
        else:
            message = MIMEText(msg)
            message['From'] = 'mentor.madness.bot@gmail.com'
            message['To'] = to_addresses
            message['Subject'] = 'Mentor meeting summary for next week'

            server.send_message(message)
            print('Sent weekly email to ' + key)

    server.quit()


def make_daily_mentor_schedules(mentor_dict):
    for key, events in mentor_dict.items():
        day = events[0].get('day')

        msg = 'Hi ' + key + ',\n\n' + \
              'Thank you for volunteering to mentor companies in the 2017 Techstars Boston cohort! Here are your ' \
              'scheduled meetings for ' + day + ':\n\n'

        event_list = bulk_event_formatter(events, for_mentors=True)
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += ''.join(event_list)

        msg += '\n\nDirections and parking instructions are attached. Please contact Ashley at (615) 719-4951' \
               ', Ty, or myself with any changes or cancellations.\n\n' \
            'Thank you,\n' \
            'Keaton'

        # Save message to a .txt file
        dirname = day.replace(' ', '_').replace('/', '_')
        txt_name = key.replace(' ', '_').strip() + '.txt'
        filename = dr.LOCAL_PATH + '/mentor_schedules/' + dirname + '/' + txt_name
        with open(filename, 'w') as file:
            file.write(msg)


def make_mentor_packet_schedules(mentor_dict):
    for key, events in mentor_dict.items():
        day = events[0].get('day')

        msg = '<h1>' + key + '</h1><br/><br/>' + \
              '<h2>Schedule - ' + day + '</h2><br/><br/>'

        event_list = bulk_event_formatter(events, for_mentors=True)
        if not event_list:
            msg += 'No meetings!'
        else:
            msg += '<br/><br/>'.join(event_list)

        # Save message to a .html file
        dirname = day.replace(' ', '_').replace('/', '_')
        txt_name = key.replace(' ', '_').strip() + '.html'
        filename = dr.LOCAL_PATH + '/mentor_packet_schedules/' + dirname + '/' + txt_name
        with open(filename, 'w') as file:
            file.write(msg)


def email_login():
    # Get login credentials from stored file
    login_file = dr.LOCAL_PATH +'mm_bot_gmail_login.txt'
    file = open(login_file, 'r')
    reader = csv.reader(file)
    row = next(reader)
    username = row[0]
    password = row[1]
    file.close()

    # Start email server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(user=username, password=password)

    return server


def bulk_event_formatter(event_dict_list, for_mentors=False):
    # First check to see if the event dict list is made up only of headers.
    # In this case, there are no actual events, so return None
    if not any(ed.get('time', False) for ed in event_dict_list):
        return []

    event_list = []
    for ed in event_dict_list:
        # Choose either mentor name or team name
        if for_mentors:
            room_subject = ed.get('company')
            if room_subject:
                room_subject = dr.names[room_subject]
        else:
            room_subject = ed.get('mentor')

        if ed.get('time') is None:
            event_list.append('\n\n' + ed['day'] + '\n')
        else:
            msg = '\n\t' + ed['time'] + ' - ' + room_subject + \
                '\n\t' + 'Room ' + ed['room_num'] + ' (' + ed['room_name'] + ')\n'
            event_list.append(msg)

    return event_list
