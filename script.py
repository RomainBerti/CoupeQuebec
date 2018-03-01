#!/usr/bin/env python
import os, sys, subprocess, csv, glob
from xlsxwriter.workbook import Workbook
from shutil import copyfile
from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
from apiclient.http import MediaFileUpload
import oauth2client
from oauth2client import client
from oauth2client import tools
import pandas

#Ajouter la copie initiale de la BD

# Get database name from arguments passed to the script
# Alternative you could set explicitly e.g. `DATABASE = 'my-access-db.mdb'`
filename = './exportedTables.csv'
DATABASE = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/GAM-2eCoupeQuebecV2.mdb'

# Get table names using mdb-tables
table_names = subprocess.Popen(['mdb-tables', '-1', DATABASE], stdout=subprocess.PIPE).communicate()[0]
tables = table_names.splitlines()


# Walk through each table and dump as CSV file using 'mdb-export'
# Replace ' ' in table names with '_' when generating CSV filename
for table in tables:
    if table == b'tblNote':
        filename = './exportedTable_tblNote.csv'
        #subprocess.call(["mdb-export", "-I", "mysql", DATABASE, table])
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)
    if table == b'tblGymnaste':
        filename = './exportedTable_tblGymnaste.csv'
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)

df_tblNotes = pandas.read_csv('./exportedTable_tblNote.csv')
df_tblGymnastes = pandas.read_csv('./exportedTable_tblGymnaste.csv')
df_Result = df_tblNotes.join(df_tblGymnastes.set_index('idGymnaste'), on='idGymnaste')
df_Result['Total'] = df_Result['sol'] + df_Result['arcon'] + df_Result['anneaux'] +
        df_Result['Saut'] + df_Result['parallele'] + df_Result['fixe'] + df_Result['Categorie']

#print (df_Result)
df_Result.to_csv('./merged.csv', columns=['Prenom','Nom', 'NomClub','sol',
                    'arcon', 'anneaux', 'Saut', 'parallele', 'fixe'], index=False)

print(tables)
# with open(filename,'r') as source:
#     rdr= csv.reader( source )
#     with open('result.csv','w') as result:
#         wtr= csv.writer( result )
#         for r in rdr:
#             wtr.writerow( (r[5], r[3], r[4], r[2], r[7], r[8]) )
# csvfile = 'result.csv'

# for csvfile in glob.glob(os.path.join('.', '*.csv')):
#     workbook = Workbook(csvfile[:-4] + '.xlsx')
#     worksheet = workbook.add_worksheet()
#     with open(csvfile, 'rt', encoding='utf8') as f:
#         reader = csv.reader(f)
#         for r, row in enumerate(reader):
#             for c, col in enumerate(row):
#                 worksheet.write(r, c, col)
#     workbook.close()






#copyfile('result.xlsx', '/Users/rpgb/Google Drive/CoupeQuebec2018/result.xlsx')
