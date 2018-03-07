# This program is to populate online website with live results during the competition
# main program to fetch marks from MS-access database and push to website

import pandas
import subprocess
import time
from shutil import copyfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# constants
PATH_DESTINATION = '/Users/rpgb/Documents/CoupeQuebec/2018'


def copy_mdb_file(PATH_DESTINATION):
    # Copy the synchronized database file to local Drive every hour
    path_to_original_database = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/2eCoupeQcGAM.mdb'
    timestamp = time.strftime('%Y%m%d_%Hh', time.localtime())
    path_to_local_database = PATH_DESTINATION + '/2eCoupeQcGAM_' + timestamp + '.mdb'

    # copy and overwrite older file to update the data but keep a backup every hour
    copyfile(path_to_original_database, path_to_local_database)

    return path_to_local_database


def get_database_info(path_to_local_database):
    # 2/ Read local mdb database and get info from tables

    # Select the desired tables and dump as CSV file using 'mdb-export'
    for table_name in ('tblNote', 'tblGymnaste'):
        filename = PATH_DESTINATION + '/exportedTable_' + table_name + '.csv'
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', path_to_local_database, table_name], stdout=f)

    # load csv in dataframe to do join on data and export result in csv with pandas
    # in the df names, I keep the original table name
    df_tblNotes = pandas.read_csv(PATH_DESTINATION + '/exportedTable_tblNote.csv')
    df_tblGymnastes = pandas.read_csv(PATH_DESTINATION + '/exportedTable_tblGymnaste.csv')

    # switch NoAffiliation and idGymnaste which are inverted
    df_tblGymnastes.columns = ['NoAffiliation', 'idGymnaste', 'Nom', 'Prenom', 'NomClub', 'Categorie',
                               'Prov', 'Age', 'Equipe']

    # Select lines for the 2eme coupe quebec CPS
    df_tblNotes = df_tblNotes[df_tblNotes['Competition'] == '2e Coupe - CPS']

    # select lines for current category
    categorie_selected = get_category(df_tblGymnastes)
    df_tblGymnastes = df_tblGymnastes[categorie_selected]

    # inner join the 2 tables on idGymnaste
    df_joinedTables = df_tblNotes.join(df_tblGymnastes.set_index('idGymnaste'),
                                       how='inner', on='idGymnaste')

    # add the column total
    df_joinedTables['Total'] = (df_joinedTables['sol'] + df_joinedTables['arcon'] +
                                df_joinedTables['anneaux'] + df_joinedTables['Saut'] +
                                df_joinedTables['parallele'] + df_joinedTables['fixe'])
    # sort dataframe by Name and firstname
    df_joinedTables.sort_values(by=['Nom', 'Prenom'], inplace=True)
    # concatenate name and firstname
    df_joinedTables['Nom'] = df_joinedTables['Prenom'] + ' ' + df_joinedTables['Nom'].str.upper()

    return df_joinedTables


def send_to_google_spreadsheet_via_api(df_joinedTables):
    # 3/ send csv file to google spreadsheet via API to be displayed
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('MyProject8533.json', scope)
    gc = gspread.authorize(credentials)
    # open an existing spreadsheet
    sheet = gc.open('ResultatsCoupeQuebec')
    ws = sheet.get_worksheet(0)
    # clear all values before updating
    ws.clear()
    # find the range of cells with the size of the dataframe
    cell_list = ws.range('A1:I' + str(df_joinedTables.shape[0]))

    cell_values = df_joinedTables[['Nom', 'NomClub', 'sol', 'arcon', 'anneaux',
                                   'Saut', 'parallele', 'fixe', 'Total']].values.flatten()

    for counter, val in enumerate(cell_values):  # gives us a tuple of an index and value
        cell_list[counter].value = val  # use the index on cell_list and the val from cell_values
    ws.update_cells(cell_list)


def get_category(df_tblGymnastes):
    # choose one of the following category by switching to True.
    saturday_morning = True
    saturday_afternoon = False
    saturday_evening = False
    sunday_morning = False
    sunday_afternoon = False
    # sunday_evening = False

    # program
    if saturday_morning:
        # Saturday from 9am to 12pm
        categorie_selected = df_tblGymnastes['Categorie'].isin(['Niveau 3 U13',
                                                               'Niveau 3 13+', 'Élite 3'])
    elif saturday_afternoon:
        # Saturday from 1pm to 5pm
        categorie_selected = df_tblGymnastes['Categorie'].isin(['Niveau 5',
                                                               'National Ouvert', 'Junior', 'Senior'])
    elif saturday_evening:
        # Saturday from 5h45pm to 8pm
        categorie_selected = df_tblGymnastes['Categorie'].isin(['Niveau 4 U13',
                                                               'Niveau 4 13+'])
    elif sunday_morning:
        # Sunday 8am to 12pm
        categorie_selected = df_tblGymnastes['Categorie'].isin(['Niveau 2A',
                                                               'Niveau 2C'])
    elif sunday_afternoon:
        # Sunday 12.30pm to 3.45pm
        categorie_selected = df_tblGymnastes['Categorie'].isin(['Niveau 2B',
                                                               'Niveau 2D'])
    # elif sunday_evening:
    #     # Sunday 4pm to 8pm
    #     categorieSelected = df_tblGymnastes['Categorie'] == 'CR'
    return categorie_selected


while True:
    local_database = copy_mdb_file(PATH_DESTINATION)
    df_joinedTables = get_database_info(local_database)
    send_to_google_spreadsheet_via_api(df_joinedTables)
    print('Website updated, waiting before next iteration... last update at ',
          time.strftime('%Hh%m', time.localtime()))
    time.sleep(30)
