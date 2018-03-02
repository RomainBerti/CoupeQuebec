# This program is to populate online website with live results during the competition
# main program to fetch marks from MS-access database and push to website

import time, os, subprocess, pandas
from shutil import copyfile
from pathlib import Path


#def main():

### 1/ Copy the sync database file to local Drive every hour
pathToOriginalDatabase = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/2eCOupeQcGAM.mdb'
pathDestination = '/Users/rpgb/Documents/CoupeQuebec/2018'
timeStamp = time.strftime("%Y%m%d_%Hh", time.localtime())
fileToCopy = pathDestination + '/2eCOupeQcGAM_' + timeStamp + '.mdb'

if  not os.path.isfile(fileToCopy):
    # file doesn't already exists
    copyfile(pathToOriginalDatabase, fileToCopy)
    print('New file copied')

### 2/ Read local mdb database and get info from tables
DATABASE = fileToCopy
categorieCourante = "Niveau 2A"

# Get table names using mdb-tables
table_names = subprocess.Popen(['mdb-tables', '-1', DATABASE], stdout=subprocess.PIPE).communicate()[0]
tables = table_names.splitlines()

# Walk through each table and dump as CSV file using 'mdb-export'
for table in tables:
    if table == b'tblNote':
        filename = pathDestination + '/exportedTable_tblNote.csv'
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)
    if table == b'tblGymnaste':
        filename = pathDestination + '/exportedTable_tblGymnaste.csv'
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)

# load csv in dataframe to do join on data and export result in csv with pandas
df_tblNotes = pandas.read_csv(pathDestination + '/exportedTable_tblNote.csv')
df_tblGymnastes = pandas.read_csv(pathDestination + '/exportedTable_tblGymnaste.csv')

# switch NoAffiliation and idGymnaste which are inverted
df_tblGymnastes.columns = ['NoAffiliation','idGymnaste', 'Nom', 'Prenom', 'NomClub', 'Categorie',
       'Prov', 'Age', 'Equipe']

#Select lines for the 2eme coupe quebec CPS
df_tblNotes = df_tblNotes[df_tblNotes['Competition'] == "2e Coupe - CPS"]
# select lines for current category

df_tblGymnastes = df_tblGymnastes[df_tblGymnastes['Categorie'] == categorieCourante]
# inner join the 2 tables on idGymnaste
df_joinedTables = df_tblNotes.join(df_tblGymnastes.set_index('idGymnaste'),
how = 'inner',on='idGymnaste')

# add the column total
df_joinedTables['Total'] = (df_joinedTables['sol'] + df_joinedTables['arcon'] +
df_joinedTables['anneaux'] + df_joinedTables['Saut'] +
df_joinedTables['parallele'] + df_joinedTables['fixe'])

#export the data to display in csv file
df_joinedTables.to_csv(pathDestination + '/merged.csv', columns=['Prenom','Nom', 'NomClub','sol',
                     'arcon', 'anneaux', 'Saut', 'parallele', 'fixe', 'Total'], index=False)
