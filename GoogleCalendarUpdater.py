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

print("Benvenuto nel programma di aggiornamento del calendario di Google Calendar")
print("Questo programma aggiorna un calendario di Google Calendar con gli eventi di un file .ods")
print("Per utilizzarlo, devi avere un file .ods con gli eventi da aggiungere al calendario")
print("Per prima cosa, devi avere un file .json con le credenziali di Google Calendar")
print("Per ottenere il file .json, segui le istruzioni su https://developers.google.com/calendar/quickstart/python")
print("Come prima cosa verranno elencati i file .ods presenti nella cartella corrente")
print("Inserisci il numero corrispondente al file che vuoi importare")
print("Successivamente verranno elencati i calendari disponibili")
print("Inserisci il numero corrispondente al calendario in cui vuoi aggiungere gli eventi")
print("Il programma proseguirà con la pulizia dei vecchi eventi e l'aggiunta dei nuovi eventi")
#variable for condition of having morning events or not
morning_events_exist = False
file_scelto=""
user_input=""
#saves list of files in the current directory
files= list(Path('./').glob("*.ods"))
#lists all the files in the current directory with .ods extension 
#and asks for the index of the file to import
if len(files)==0:
    print("Non ci sono files .ods nella cartella corrente")
    input("Premi invio per uscire")
    exit()
while file_scelto=="" and (user_input!="s" or user_input!="S"):
    #resets user_input and file_scelto to avoid errors
    user_input=""
    file_scelto=""
    print("Quale tra questi files vuoi importare? (inserisci il numero corrispondente)")
    for i, latest in enumerate(files):
        print(f"{i}: {latest}")
    try:
        #asks for the index of the file to import
        file_da_cercare=int(input())
        #checks if the index is valid
        if file_da_cercare < len(files):
            #asks if the user wants to continue
            print("Hai scelto il file ",files[file_da_cercare], " vuoi continuare? (s/n)")
            user_input=input()
            #if the user wants to continue, the file is saved in file_scelto
            if user_input=="s" or user_input=="S":
                file_scelto=files[file_da_cercare]
        else: 
            print("Il numero non è valido, riprova")
    #if the user doesn't insert a number, the program asks again
    except ValueError:
        print("Il dato inserito non è valido, riprova")
#latest=max([f for f in files], key=lambda item: item.stat().st_ctime)

#saves list of calendars available
calendars=list(list_of_calendars())


calendar_id=""
#lists all the calendars available and asks for the index of the calendar to update
while calendar_id=="" and (user_input!="s" or user_input!="S"):
    #resets user_input and calendar_id to avoid errors
    user_input=""
    calendar_id=""
    print("Quale tra questi calendari vuoi aggiornare? (inserisci il numero corrispondente)")
    for i, calendar in enumerate(calendars):
        print(f"{i}: {calendar['summary']}")
    #asks for the index of the calendar to update
    try:
        calendar_index=int(input())
        #checks if the index is valid
        if calendar_index < len(calendars):
            calendario=calendars[calendar_index]
            #asks if the user wants to continue
            print("Hai scelto il calendario ",calendario['summary'], " tutti gli eventi presenti verranno cancellati, vuoi continuare? (s/n)")
            user_input=input()
            #if the user wants to continue, the calendar['id'] is saved in calendar_id
            if user_input=="s" or user_input=="S":
                #finds the calendar id of the calendar with the summary of the calendar chosen
                #as calendar_id and summary are not the same, i have to search for the summary
                for calendar in calendars:
                    summary = calendar['summary']
                    if calendario['summary']==summary:
                        calendar_id=calendar['id']
        else:
            #if the index is not valid, the program asks again
            print("Non hai scelto un calendario valido, riprova")
    #if the user doesn't insert a number, the program asks again
    except ValueError:
        print("Il dato inserito non è valido, riprova")

# Delete all events in the calendar with the name of calendar_name
service = get_calendar_service()
#creates a batch request to delete all the events in the calendar
batch_delete = service.new_batch_http_request()
try:
    events_result = service.events().list(calendarId=calendar_id, maxResults=2499).execute()
    k=0
    n_event_to_delete=int(len(events_result.get('items', [])))
    print (n_event_to_delete," eventi da eliminare nel calendario ",calendario['summary'])
    for event in events_result.get('items', []):
        batch_delete.add(service.events().delete(calendarId=calendar_id, eventId=event['id']))
        #for debug
        #print ("Event deleted: ", event['summary'],event['description'],event['start'])
        
        #print progress
        k+=1
        if k%10==0:
            print("Eventi eliminati: ",k,"/",n_event_to_delete," ",int(k*100/n_event_to_delete), '%')
    #print if all events have been deleted or not
    batch_delete.execute()
    if k<n_event_to_delete:
        print("ERRORE, non sono stati eliminati tutti gli eventi!")
    else:
        print("Tutti gli eventi sono stati eliminati!")
    print(k," eventi eliminati in totale nel calendario ",calendario['summary'])    
except errors.HttpError:
    print("Fallito a cancellare gli eventi nel calendario ",calendario['summary'])

# read odf file and convert into a dataframe object
df = pd.DataFrame(pd.read_excel (file_scelto, engine="odf"))

#removes all empty and useless columns
df=df[["Giorno","Orario mattino","Docente","Unità formativa","Orario pomeriggo","Docente.1","Unità formativa.1"]]
df.dropna(axis=1,how='all',inplace=True )

#check if there are events in the morning
if 'Orario mattino' in df.columns:
    morning_events_exist = True

#fills dates that have 2 events the same afternoon or morning
s = df["Giorno"].eq("")
df.loc[s, "Giorno"] = np.nan
df["Giorno"].ffill(inplace=True)

#managing morning events
if morning_events_exist == True:
    #the easy way to manage morning events is to create a separate dataset and then merge them
    df2=df[['Giorno','Orario mattino','Docente','Unità formativa']] 
    
    #removes the morning columns from the afternoon df
    df.drop(columns=["Orario mattino","Docente","Unità formativa"],inplace=True)
    
    #now i do the formatting for the morning dataframe
    df2['Orario mattino'] = df2['Orario mattino'].str.replace('–','-',regex=False)
    df2['Orario mattino'] = df2['Orario mattino'].str.replace('.',':',regex=False)
    df2['Orario mattino'] = df2['Orario mattino'].str.replace(',',':',regex=False)
    
    df2[['Start Time','End Time']]=df2['Orario mattino'].str.split(pat='-',n=1,expand=True,regex=False)
    
    df2.drop(columns=["Orario mattino"],inplace=True)
    
    df2.rename(columns={"Docente":"Description"},inplace=True)
    df2.rename(columns={"Unità formativa":"Subject"},inplace=True)
    df2.rename(columns={"Giorno":"Start Date"},inplace=True)
    
    #gets df2 ready for merging
    df2=df2[['Subject','Start Date','Start Time','End Time','Description']]

#start formatting afternoon hours
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('–','-',regex=False)
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('.',':',regex=False)
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace(',',':',regex=False)

#splitting in start tame and end time
df[['Start Time','End Time']]=df['Orario pomeriggo'].str.split(pat='-',n=1,expand=True,regex=False)

#renaming columns for calendar csv template
df.rename(columns={"Docente.1":"Description"},inplace=True)
df.rename(columns={"Unità formativa.1":"Subject"},inplace=True)
df.rename(columns={"Giorno":"Start Date"},inplace=True)

#drops all columns not needed in the calendar csv
if morning_events_exist == True:
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

# Create the batch request to send to the API to create events
batch_insert = service.new_batch_http_request()
try:
    # Itereate over the dataframe and create events
    
    j=0
    numero_eventi_da_creare=int(len(df.index))
    print("Creazione di ",numero_eventi_da_creare," eventi nel calendario ",calendario['summary'])
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
        # Add the event to the batch
        batch_insert.add(service.events().insert(calendarId=calendar_id, body=event))
        #print the progress
        j+=1
        if j%10==0:
            print("Eventi creati: ",j,"/",numero_eventi_da_creare," ",int(j*100/numero_eventi_da_creare), '%')
    # Execute the batch 
    batch_insert.execute()
    print(j , " eventi creati in totale nel calendario ",calendario['summary'])
except errors.HttpError:
        print("Fallito a creare gli eventi nel calendario ",calendario['summary'])
        
#prints the summary of the import
print("Importazione completata, riassunto:")
print("dal file: ",file_scelto )
print("calendario aggiornato: ",calendario['summary'])
print("eventi eliminati: ",k)
print("eventi creati: ",j)
input("Premi un tasto per uscire")