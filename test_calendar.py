#note filter to see all columns that are enterly emtpy df.isnull().all()


#importing pandas as pd
import pandas as pd
import numpy as np
import time
from pathlib import Path
#finds last calendar
files= Path('./').glob("Calendario II anno Imola al*")
latest=max([f for f in files], key=lambda item: item.stat().st_ctime)
# read odf file and convert 
# into a dataframe object
df = pd.DataFrame(pd.read_excel (latest, engine="odf"))
#removes all empty columns
df.dropna(axis=1,how='all',inplace=True )
#fills dates that have 2 events the same afternoon
s = df["Giorno"].eq("")
df.loc[s, "Giorno"] = np.nan
df["Giorno"].ffill(inplace=True)
#start formatting hours
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('–','-')
df['Orario pomeriggo'] = df['Orario pomeriggo'].str.replace('.',':')
#splitting in start tame and end time
df[['Start Time','End Time']]=df['Orario pomeriggo'].str.split('-',1,expand=True)
#drops all columns not needed in the calendar csv
df.drop(columns=["Unnamed: 9","Unnamed: 10","Unnamed: 11","tot ore","Orario pomeriggo","Unità formativa","giorno settimana"],inplace=True)
#renaming columns for calendar csv template
df.rename(columns={"Docente.1":"Description"},inplace=True)
df.rename(columns={"Unità formativa.1":"Subject"},inplace=True)
df.rename(columns={"Giorno":"Start Date"},inplace=True)
#removes rows with no subject or start date
df=df.dropna(subset=['Start Date','Subject'])
#reformatting date in same way as csv template and adds a end date
df['Start Date'] = pd.to_datetime(df['Start Date'],errors='coerce')
df['Start Date']=df['Start Date'].dt.strftime('%m/%d/%Y')
df['End Date']=df['Start Date']
#add needed columns
df[["All Day Event","Location"]]=""
df[["Private","All Day Event"]]=False
#reorder columns like the template
df = df[['Subject','Start Date','Start Time','End Date','End Time','All Day Event','Description','Location','Private']]

#saves it in csv with timestamp
snapshotdate = time.strftime('%d-%m-%Y')
filenamewithdate="Calendario Trasformato "+snapshotdate+".csv"
df.to_csv (filenamewithdate, index = None, header=True)