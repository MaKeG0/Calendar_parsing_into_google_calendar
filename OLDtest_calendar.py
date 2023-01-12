#deprecated
#importing pandas as pd
import pandas as pd
#importing numpy as pd
import numpy as np
import time
from pathlib import Path

#variable for condition of having morning events or not
morning = False

#finds last calendar

files= Path('./').glob("*.ods")
latest=max([f for f in files], key=lambda item: item.stat().st_ctime)

# read odf file and convert into a dataframe object
df = pd.DataFrame(pd.read_excel (latest, engine="odf"))

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

#reformatting date in same way as csv template and adds a end date
df['Start Date'] = pd.to_datetime(df['Start Date'],errors='coerce')
df['Start Date']=df['Start Date'].dt.strftime('%m/%d/%Y')

# df['End Date']=df['Start Date'] not necessary
#add needed columns
df[["Private","All Day Event"]]=False
#reorder columns like the template
df = df[['Subject','Start Date','Start Time','End Time','All Day Event','Description','Private']]

#saves it in csv with timestamp
snapshotdate = time.strftime('%d-%m-%Y')
filenamewithdate="Calendario Trasformato "+snapshotdate+".csv"
df.to_csv (filenamewithdate, index = None, header=True)
