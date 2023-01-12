#importing pandas as pd
import pandas as pd
#importing numpy as np
import numpy as np
#import pathlib to find the latest file
from pathlib import Path
#importing calendar api functions from cal_setup.py and list_calendars.py
from cal_setup import get_calendar_service
from list_calendars import list_of_calendars
#importing googleapiclient errors to handle errors in the calendar api
from googleapiclient import errors


#variable for condition of having morning events or not
morning = False
file_scelto=""
ok=""
#finds last calendar
files= list(Path('./').glob("*.ods"))
while file_scelto=="" and (ok!="s" or ok!="S"):
    ok=""
    file_scelto=""
    print("Quale tra questi files vuoi importare? (inserisci il numero corrispondente)")
    for i, latest in enumerate(files):
        print(f"{i}: {latest}")
    try:
        file_da_cercare=int(input())
        if file_da_cercare < len(files):
            print("Hai scelto il file ",files[file_da_cercare], " vuoi continuare? (s/n)")
            ok=input()
            if ok=="s" or ok=="S":
                file_scelto=files[file_da_cercare]
        else: 
            print("Il numero non è valido, riprova")
    except ValueError:
        print("Il dato inserito non è valido, riprova")
#latest=max([f for f in files], key=lambda item: item.stat().st_ctime)

#saves list of calendars available
calendars=list(list_of_calendars())

#lists all the calendars available
calendar_id=""
while calendar_id=="" and (ok!="s" or ok!="S"):
    ok=""
    calendar_id=""
    print("Quale tra questi calendari vuoi aggiornare? (inserisci il numero corrispondente)")
    for i, calendar in enumerate(calendars):
        print(f"{i}: {calendar['summary']}")
        
    #asks for the index of the calendar to update
    try:
        calendar_index=int(input())
        if calendar_index < len(calendars):
            calendario=calendars[calendar_index]
            print("Hai scelto il calendario ",calendario['summary'], " vuoi continuare? (s/n)")
            ok=input()
            if ok=="s" or ok=="S":
                for calendar in calendars:
                    summary = calendar['summary']
                    if calendario['summary']==summary:
                        calendar_id=calendar['id']
        else:
            print("Non hai scelto un calendario valido, riprova")
    except ValueError:
        print("Il dato inserito non è valido, riprova")

# Delete all events in the calendar with the name of calendar_name
service = get_calendar_service()
try:
    events_result = service.events().list(calendarId=calendar_id).execute()
    for event in events_result.get('items', []):
        service.events().delete(calendarId=calendar_id, eventId=event['id']).execute()
        print ("Event deleted: ", event['summary'],event['description'],event['start'])
    print("Eventi eliminati nel calendario ",calendario['summary'])    
except errors.HttpError:
    print("Fallito a cancellare gli eventi nel calendario ",calendario['summary'])

# read odf file and convert into a dataframe object
df = pd.DataFrame(pd.read_excel (file_scelto, engine="odf"))

#removes all empty and useless columns
df=df[["Giorno","Orario mattino","Docente","Unità formativa","Orario pomeriggo","Docente.1","Unità formativa.1"]]
df.dropna(axis=1,how='all',inplace=True )

#check if there are events in the morning
if 'Orario mattino' in df.columns:
    morning = True

#fills dates that have 2 events the same afternoon or morning
s = df["Giorno"].eq("")
df.loc[s, "Giorno"] = np.nan
df["Giorno"].ffill(inplace=True)

#managing morning events
if morning == True:
    #the easy way to manage morning events is to create a separate dataset and then merge them
    df2=df[['Giorno','Orario mattino','Docente','Unità formativa']] 
    
    #removes the morning columns from the afternoon df
    df.drop(columns=["Orario mattino","Docente","Unità formativa"],inplace=True)
    
    #now i do the formatting for the morning dataframe
    df2['Orario mattino'] = df2['Orario mattino'].str.replace('–','-')
    df2['Orario mattino'] = df2['Orario mattino'].str.replace('.',':')
    
    df2[['Start Time','End Time']]=df2['Orario mattino'].str.split('-',1,expand=True)
    
    df2.drop(columns=["Orario mattino"],inplace=True)
    
    df2.rename(columns={"Docente":"Description"},inplace=True)
    df2.rename(columns={"Unità formativa":"Subject"},inplace=True)
    df2.rename(columns={"Giorno":"Start Date"},inplace=True)
    
    #gets df2 ready for merging
    df2=df2[['Subject','Start Date','Start Time','End Time','Description']]

#start formatting afternoon hours
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('–','-')
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('.',':')

#splitting in start tame and end time
df[['Start Time','End Time']]=df['Orario pomeriggo'].str.split('-',1,expand=True)

#renaming columns for calendar csv template
df.rename(columns={"Docente.1":"Description"},inplace=True)
df.rename(columns={"Unità formativa.1":"Subject"},inplace=True)
df.rename(columns={"Giorno":"Start Date"},inplace=True)

#drops all columns not needed in the calendar csv
if morning == True:
    df.drop(columns=["Orario pomeriggo"],inplace=True)

    #gets df ready for merging
    df=df[['Subject','Start Date','Start Time','End Time','Description']]
    
    #if there are morning events it merges the 2 dataframe before the last formatting
    df=pd.concat([df,df2],ignore_index=True)
else:
    df.drop(columns=["Orario pomeriggo","Unità formativa"],inplace=True)

#removes rows with no subject or start date or start time
df=df.dropna(subset=['Start Date','Subject','Start Time'])

#reformatting date in same way as google api template
df['Start Date'] = pd.to_datetime(df['Start Date'],errors='coerce')
df['Start Date']=df['Start Date'].dt.strftime('%Y-%m-%d')
df['Start Time']=df['Start Time']
df['End Time']=df['End Time']

#reorder columns like the template
df = df[['Subject','Start Date','Start Time','End Time','Description']]


def get_color_id(summary):
    # Use a hash function to generate a colorId
    color_id = (hash(summary) % 10)+1
    return str(color_id)

try:
    # Itereate over the dataframe and create events
    for index, row in df.iterrows():
        # Create the event json to be sent to the API
        event = {
            'summary': row['Subject'],
            'description': row['Description'],
            # match the color of the event with the color of the calendar
            'colorId':  get_color_id(row['Subject']),
            'start': {
                'dateTime': row['Start Date'].strip() + 'T' + row['Start Time'].strip()+':00',
                'timeZone': 'Europe/Rome'
            },
            'end': {
                'dateTime': row['Start Date'].strip() + 'T' + row['End Time'].strip()+':00',
                'timeZone': 'Europe/Rome'
            },
        }
        #send the event to the API
        service.events().insert(calendarId=calendar_id, body=event).execute()
        i+=1
        if i%10==0:
            print("Eventi creati: ",i)
    print("Eventi creati nel calendario ",calendario['summary'])
except errors.HttpError:
        print("Fallito a creare gli eventi nel calendario ",calendario['summary'])