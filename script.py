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

#Ajouter la copie initiale de la BD

# Get database name from arguments passed to the script
# Alternative you could set explicitly e.g. `DATABASE = 'my-access-db.mdb'`
filename = './exportedTables.csv'
DATABASE = '/Users/rpgb/Dropbox/CoupeQuebec/2018/WebSitePythonParsing/BDD-live/GAM-2eCoupeQuebecV2.mdb'

# Get table names using mdb-tables
table_names = subprocess.Popen(['mdb-tables', '-1', DATABASE], stdout=subprocess.PIPE).communicate()[0]
tables = table_names.splitlines()


# Walk through each table and dump as CSV file using 'mdb-export'
# Replace ' ' in table names with '_' when generating CSV filename
for table in tables:
    if table == b'tblNote':
        subprocess.call(["mdb-export", "-I", "mysql", DATABASE, table])
        #print('Exporting ' + table)
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', DATABASE, table], stdout=f)
print(tables)
with open(filename,'r') as source:
    rdr= csv.reader( source )
    with open('result.csv','w') as result:
        wtr= csv.writer( result )
        for r in rdr:
            wtr.writerow( (r[5], r[3], r[4], r[2], r[7], r[8]) )
csvfile = 'result.csv'

for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet()
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()





file_metadata = {'name': 'My Report','mimeType': 'application/vnd.google-apps.spreadsheet'}
#Now create the media file upload object and tell it what file to upload,
media = MediaFileUpload('result.csv', mimetype='text/csv',resumable=True)
file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
print('File ID: %s' % file.get('id'))

#copyfile('result.xlsx', '/Users/rpgb/Google Drive/CoupeQuebec2018/result.xlsx')
