import os
import discord
import chat_exporter
import re
import sys
import logging

from discord.ext import tasks
from dotenv import load_dotenv
from bs4 import BeautifulSoup
from datetime import datetime
from dateutil import parser
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

logger = logging.getLogger(__name__)
handler = logging.StreamHandler(stream=sys.stdout)
logger.addHandler(handler)

load_dotenv()
discord_bot_token = os.getenv('DISCORD_BOT_TOKEN')
discord_server_id = os.getenv('DISCORD_SERVER_ID')
discord_channel_id = os.getenv('DISCORD_CHANNEL_ID')
discord_checked_messages = os.getenv('DISCORD_CHECKED_MESSAGES')
discord_check_frequency = os.getenv('DISCORD_CHECK_FREQUENCY_S')
google_client_scope = os.getenv('GOOGLE_CLIENT_SCOPE')
google_spreadsheet_id = os.getenv('GOOGLE_SPREADSHEET_ID')
google_worksheet_name = os.getenv('GOOGLE_WORKSHEET_NAME')
google_worksheet_start_row = int(os.getenv('GOOGLE_WORKSHEET_START_ROW'))
google_worksheet_max_rows = int(os.getenv('GOOGLE_WORKSHEET_MAX_ROWS'))

client = discord.Client()
channel = None
creds = None

# This function handles exceptions
def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

# This function gets called when the bot is connected to discord
@client.event
async def on_ready():
    global channel
    
    print('\033[92m' + f'\nConnected to discord as user "{client.user}"')
    
    try:
        channel = await client.fetch_channel(discord_channel_id)
    except:
        print('\033[91m' + f'Failed to find channel with ID {discord_channel_id}')
        return
    else:
        print('\033[92m' + f'Found channel: "{channel.name}"')
        update_event.start()
        

# The function gets called to update the event data in the spreadsheet
@tasks.loop(seconds=int(discord_check_frequency))
async def update_event():
    global channel
    global creds
    
    try:
        history = await channel.history(limit=int(discord_checked_messages)).flatten()
        print('\n')
    except:
        print('\n')
        print('\033[91m' + 'Failed to find channel history')
        return
    
    try:
        html_text = await chat_exporter.raw_export(channel, history, set_timezone='UTC')
        html_object = BeautifulSoup(html_text, features='html.parser')
    except:
        print('\033[91m' + 'Failed to convert channel history into an object')
        return
    
    try:
        messages_list = html_object.find_all('div', class_='chatlog__messages')
        reversed_messages_list = list(reversed(messages_list))
    except:
        print('\033[91m' + 'Failed to get messages list')
        return
        
    try:
        if (os.path.exists('token.json')):
            print('\033[92m' + 'Getting existing credentials from file')
            creds = Credentials.from_authorized_user_file('token.json', [google_client_scope])
        # Check if there are valid credentials
        if not (creds) or not (creds.valid):
            refreshed = False
            if (creds) and (creds.expired) and (creds.refresh_token):
                # If the access token expired, use the refresh token to get another one
                print('\033[92m' + 'Refreshing access token using refresh token')
                try:
                    creds.refresh(Request())
                    refreshed = True
                except:
                    #Do nothing
                    pass
            if not (refreshed):
                # We haven't gotten an access and refresh token yet, so get them using the client credentials
                print('\033[92m' + 'Getting new credentials...')
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', [google_client_scope])
                creds = flow.run_local_server(port=8080, open_browser=True)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                print('\033[92m' + 'Writing new credentials to file')
                token.write(creds.to_json())
    except:
        print('\033[91m' + 'Failed to get credentials')
        return

    try:
        spreadsheets = build('sheets', 'v4', credentials=creds).spreadsheets()
    except:
        print('\033[91m' + 'Failed to find spreadsheet')
        return
        
    valid_messages = 0
    values_array = []
    
    for (message) in (reversed_messages_list):
        try:
            event_title = message.find('span', class_='markdown').text
        except:
            continue
        else:
            print('\033[92m' + 'Found event title: "' + event_title + '"')
    
            discord_timestamp_format = message.find('span', class_='chatlog__timestamp').text
            if not (discord_timestamp_format):
                print('\033[91m' + 'Failed to find discord timestamp')
                #print(f'\n{message}')
                continue
            else:
                discord_timestamp = parser.parse(discord_timestamp_format).utcnow().replace(microsecond=0).isoformat()
                print('\033[92m' + 'Found discord timestamp: "' + discord_timestamp + '"')
                
            embed_description = message.find('div', class_='chatlog__embed-description')
            if not (embed_description):
                event_description = ''
            else:
                event_description = embed_description.find('span', class_='markdown preserve-whitespace').text
                if not (event_description):
                    print('\033[91m' + 'Failed to find event description')
                    #print(f'\n{message}')
                    continue
                    
            print('\033[92m' + 'Found event description: "' + event_description + '"')
                
            creator_text = message.find('span', class_='chatlog__embed-footer-text').text
            if not (creator_text):
                print('\033[91m' + 'Failed to find event creator')
                continue
            else:
                creator_text = creator_text.replace('Created by ', '')
                creator_text = re.split(' • ', creator_text)
                event_creator = creator_text[0]
                if (len(creator_text) > 1):
                    event_repeat = creator_text[1]
                else:
                    event_repeat = 'No'
                    
            print('\033[92m' + 'Found event creator: "' + event_creator + '"')
            print('\033[92m' + 'Found event repeat: "' + event_repeat + '"')
            
            
            embed_fields = message.find('div', class_='chatlog__embed-fields')
            try:
                embed_field_list = embed_fields.find_all('span')
                span_text = str(embed_field_list[1])
                split_text = re.split('/>|</|>|<| - |br|t:|:f|:r|:t| ', span_text)
                split_text = list(filter(None, split_text))
                ts = int(split_text[3])
                event_start_timestamp = datetime.utcfromtimestamp(ts).isoformat()
                if (split_text[4].isnumeric()):
                    ts = int(split_text[4])
                    event_end_timestamp = datetime.utcfromtimestamp(ts).isoformat()
                else:
                    event_end_timestamp = ''
            except:
                print('\033[91m' + 'Failed to find event timestamps')
                #print(f'\n{message}')
                continue
            
            print('\033[92m' + 'Found event start time: "' + event_start_timestamp + '"')
            print('\033[92m' + 'Found event end time: "' + event_end_timestamp + '"')
            
            values_array.append([event_start_timestamp, event_end_timestamp, event_repeat, event_creator, discord_timestamp, event_title, event_description])
            valid_messages = valid_messages + 1
                    
                    
    update_range = google_worksheet_name + '!' + str(google_worksheet_start_row) + ':' + str(google_worksheet_start_row + google_worksheet_max_rows)
    
    value_range = {
                "majorDimension": "ROWS",
                'values': values_array
            }
            
    if (valid_messages < google_worksheet_max_rows):
        for (x) in range(valid_messages, google_worksheet_max_rows):
            values_array.append(['', '', '', '', '', '', ''])
        
    try:
        response = spreadsheets.values().update(spreadsheetId=google_spreadsheet_id, range=update_range, valueInputOption="USER_ENTERED", body=value_range).execute()
        print('\033[92m' + 'Finished updating worksheet')
    except Exception as e:
        print('\033[91m' + 'Failed to update worksheet')
        return
                
    print('\033[92m' + 'Number of rows with event data: ' + str(valid_messages))
    print('\033[92m' + 'Number of rows with empty data: ' + str(google_worksheet_max_rows - valid_messages))
        
        
sys.excepthook = handle_exception

client.run(discord_bot_token)
