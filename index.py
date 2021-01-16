from __future__ import print_function
import datetime
import pickle
import os.path
import pandas as pd
import subprocess
import pyautogui
import time
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.s
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

# Global variables
events_df = pd.read_csv('data.csv')   # Create DataFrame to store event info
day_type = input("A/B/X Day? ")

def calendar_api():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Set datetime range for today only in correct format

    #day_start = str(datetime.date.today()) + 'T' + '00:00:0.0Z'
    now = datetime.datetime.utcnow().isoformat() + 'Z'
    day_end = str(datetime.date.today()) + 'T' + '23:59:59.0Z'

    # Get events from School Activities calendar for today
    events_result = service.events().list(calendarId='leanderisd.org_c684tcob9a0a8eufg3ctq2bu2o@group.calendar.google.com', timeMin=now, timeMax=day_end, maxResults=20, singleEvents=True, orderBy='startTime').execute()
    
    events = events_result.get('items', [])
    
    if not events:
        print('No upcoming events found.')
        sys.exit()
    else:
        return events

def sign_in(meetingid,pswd):

    # opens up zoom app
    subprocess.call(["C:/Users/Asha_karmakar/AppData/Roaming/Zoom/bin/Zoom.exe"])

    time.sleep(2)

    # clicks on join button
    join_btn = pyautogui.locateCenterOnScreen('images/join_button.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()

    # enters meeting id
    time.sleep(2)
    pyautogui.write(meetingid)

    # to disable audio and video
    '''
    check_box = pyautogui.locateAllOnScreen('images/check_box.png')
    for box in check_box:
        pyautogui.moveTo(box)
        pyautogui.click()
        time.sleep(2)
    '''

    # clicks on join button
    join_btn = pyautogui.locateCenterOnScreen('images/join_button_2.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()

    # enters password
    time.sleep(2)
    pyautogui.write(pswd)

    # clicks on join button
    join_btn = pyautogui.locateCenterOnScreen('images/join_button_3.png')
    pyautogui.moveTo(join_btn)
    pyautogui.click()

def setup(events):
    count = 0
    for event in events:
        try:
            #print(event['summary'])
            start_time = event['start'].get('dateTime', event['start'].get('date'))   
            start_time = start_time[start_time.find('T')+1: start_time.index('-06:00')]

            if day_type == 'A':
                period = int(str(event['summary']).split()[1][0])
            elif day_type == 'B':
                period = int(str(event['summary']).split()[1][2])
                    
            #print (period)

            try:
                ind = events_df.index[events_df['period'] == period][0]
                events_df.at[ind, 'time'] = start_time
                count += 1
            except:
                pass
        except:
            print ("Error in events list")
            sys.exit()

    print (events_df)
    return count

def main():

    if (day_type == 'X'):
        sys.exit()
        
    events = calendar_api()
    numEvents = setup(events)
    
    print (numEvents, " events scheduled")

    while numEvents > 0:  
        now = datetime.datetime.now() # Get current time
        now = str((now - datetime.timedelta(microseconds=now.microsecond)).time()) # Round to nearest second and convert to string to match DataFrame time format
        #now = '12:31:00' #<-- Sample time for testing

        if now in events_df['time'].tolist(): # Check if there is any event scheduled at this time

            ind = events_df.index[events_df['time'] == now][0] # Get index of event in df
            
            m_period = str(events_df.at[ind, 'period']) # Extract event info from df
            m_id = str(events_df.at[ind, 'meeting_id'])
            m_pwd = str(events_df.at[ind, 'meeting_password'])

            print (m_id, m_pwd)
            print ("Joining Zoom for Period ", m_period)

            try:
                sign_in(m_id,m_pwd) # Log into Zoom meeting
                #sign_in('4114270424','w57w5v') <-- Sample Meeting
            except:
                print ("Unable to sign into Zoom meeting.")

            numEvents -= 1
            print (numEvents, " remaining")

if __name__ == '__main__':
    main()