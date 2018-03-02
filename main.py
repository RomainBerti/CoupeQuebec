# This program is to populate online website with live results during the competition
# main program to fetch marks from MS-access database and push to website

import time, os, subprocess, pandas#, csv
from shutil import copyfile
from pathlib import Path
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.ns import _ns
from functions import getcategory

def main():

    ### 1/ Copy the sync database file to local Drive every hour
    pathToOriginalDatabase = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/2eCOupeQcGAM.mdb'
    pathDestination = '/Users/rpgb/Documents/CoupeQuebec/2018'
    timeStamp = time.strftime("%Y%m%d_%Hh", time.localtime())
    fileToCopy = pathDestination + '/2eCOupeQcGAM_' + timeStamp + '.mdb'
    # copy and overwrite older file to update the data
    copyfile(pathToOriginalDatabase, fileToCopy)

    ### 2/ Read local mdb database and get info from tables
    DATABASE = fileToCopy

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

    #condition = ((df_tblGymnastes['Categorie'] == "Niveau 2A") | (df_tblGymnastes['Categorie'] == "National Ouvert") )
    categorieSelected = getcategory(df_tblGymnastes)
    df_tblGymnastes = df_tblGymnastes[categorieSelected]

    # inner join the 2 tables on idGymnaste
    df_joinedTables = df_tblNotes.join(df_tblGymnastes.set_index('idGymnaste'),
    how = 'inner',on='idGymnaste')

    # add the column total
    df_joinedTables['Total'] = (df_joinedTables['sol'] + df_joinedTables['arcon'] +
    df_joinedTables['anneaux'] + df_joinedTables['Saut'] +
    df_joinedTables['parallele'] + df_joinedTables['fixe'])
    df_joinedTables.sort_values(by=['Nom', 'Prenom'], inplace=True)
    df_joinedTables['Nom'] = df_joinedTables['Prenom'] + ' ' + df_joinedTables['Nom'].str.upper()

    ### 3/ send csv file to google spreadsheet via API to be displayed
    scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('MyProject8533.json', scope)
    gc = gspread.authorize(credentials)
    #open an existing spreadsheet
    sheet = gc.open('ResultatsCoupeQuebec')
    ws = sheet.get_worksheet(0)
    #clear all values before updating
    ws.clear()
    #find the range of cells with the size of the dataframe
    cell_list = ws.range('A1:I' + str(df_joinedTables.shape[0]))

    cell_values = df_joinedTables[['Nom', 'NomClub','sol','arcon', 'anneaux',
    'Saut', 'parallele', 'fixe', 'Total']].values.flatten()

    for counter, val in enumerate(cell_values):  #gives us a tuple of an index and value
        cell_list[counter].value = val    #use the index on cell_list and the val from cell_values
    ws.update_cells(cell_list)

#while True:
main()
    # print('Website updated, waiting before next iteration... last update at ',
    # time.strftime("%Hh%m", time.localtime()))
    # time.sleep(30)
