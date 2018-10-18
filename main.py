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
PRIVATE_KEY_JSON = '/Users/rpgb/Documents/CoupeQuebec/2018/MyProject8533.json'
ORIGINAL_DATABASE_PATH = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/2eCoupeQcGAM.mdb'

def copy_mdb_file(ORIGINAL_DATABASE_PATH, PATH_DESTINATION):
    """
    :param PATH_DESTINATION: location to save the database locally
    :param ORIGINAL_DATABASE_PATH: fullpath to DB to save
    :return: path_to_local_database: full path with DB local filename that changes every hour
    """
    customed_timestamp = time.strftime('%Y%m%d_%Hh', time.localtime())
    path_to_local_database = PATH_DESTINATION + '/2eCoupeQcGAM_' + customed_timestamp + '.mdb'

    # copy and overwrite older file to update the data but keep a backup every hour
    copyfile(ORIGINAL_DATABASE_PATH, path_to_local_database)

    return path_to_local_database


def get_info_to_be_displayed_from_database(path_to_local_database):
    """
    :param path_to_local_database: full path with DB local filename that changes every hour
    :return: df_tbl_notes_with_gymnastes: contains the notes with the gymnastes names to be displayed
    """

    # Select the desired tables and dump as CSV file using 'mdb-export'
    for table_name in ('tblNote', 'tblGymnaste'):
        filename = PATH_DESTINATION + '/exportedTable_' + table_name + '.csv'
        with open(filename, 'wb') as f:
            subprocess.check_call(['mdb-export', path_to_local_database, table_name], stdout=f)

    # load csv in dataframe to do join on data and export result in csv with pandas
    # in the df names, image keep the original table name
    df_tbl_notes = pandas.read_csv(PATH_DESTINATION + '/exportedTable_tblNote.csv')
    df_tbl_gymnastes = pandas.read_csv(PATH_DESTINATION + '/exportedTable_tblGymnaste.csv')

    # switch NoAffiliation and idGymnaste which are inverted
    df_tbl_gymnastes.columns = ['NoAffiliation', 'idGymnaste', 'Nom', 'Prenom', 'NomClub', 'Categorie',
                                'Prov', 'Age', 'Equipe']

    # Select lines for the 2eme coupe quebec CPS
    df_tbl_notes = df_tbl_notes[df_tbl_notes['Competition'] == '2e Coupe - CPS']

    # select lines for current category
    categorie_selected = get_category(df_tbl_gymnastes)
    df_tbl_gymnastes = df_tbl_gymnastes[categorie_selected]

    # inner join the 2 tables on idGymnaste
    df_tbl_notes_with_gymnastes = df_tbl_notes.join(df_tbl_gymnastes.set_index('idGymnaste'),
                                                    how='inner', on='idGymnaste')

    # add the column total
    df_tbl_notes_with_gymnastes['Total'] = (df_tbl_notes_with_gymnastes['sol'] + df_tbl_notes_with_gymnastes['arcon'] +
                                            df_tbl_notes_with_gymnastes['anneaux'] + df_tbl_notes_with_gymnastes['Saut'] +
                                            df_tbl_notes_with_gymnastes['parallele'] + df_tbl_notes_with_gymnastes['fixe'])
    # sort dataframe by Name and firstname
    df_tbl_notes_with_gymnastes.sort_values(by=['Nom', 'Prenom'], inplace=True)
    # concatenate name and firstname
    df_tbl_notes_with_gymnastes['Nom'] = df_tbl_notes_with_gymnastes['Prenom'] + ' ' + df_tbl_notes_with_gymnastes['Nom'].str.upper()

    return df_tbl_notes_with_gymnastes


def send_to_google_spreadsheet_via_api(PRIVATE_KEY_JSON, df_tbl_notes_with_gymnastes):
    """
    :param PRIVATE_KEY_JSON: path to the private key for the authentification for the API
    :param df_tbl_notes_with_gymnastes:  contains the notes with the gymnastes names to be displayed
    :return: none
    """
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(PRIVATE_KEY_JSON, scope)
    gc = gspread.authorize(credentials)
    # open an existing spreadsheet
    sheet = gc.open('ResultatsCoupeQuebec')
    ws = sheet.get_worksheet(0)
    # clear all values before updating
    ws.clear()
    # find the range of cells with the size of the dataframe
    cell_list = ws.range('A1:image' + str(df_tbl_notes_with_gymnastes.shape[0]))

    cell_values = df_tbl_notes_with_gymnastes[['Nom', 'NomClub', 'sol', 'arcon', 'anneaux',
                                   'Saut', 'parallele', 'fixe', 'Total']].values.flatten()

    for counter, val in enumerate(cell_values): 
        cell_list[counter].value = val  # use the index on cell_list and the val from cell_values
    ws.update_cells(cell_list)


def get_category(df_tbl_gymnastes):
    """ select a category by manually switching the boolean variables"""
    saturday_morning = True
    saturday_afternoon = False
    saturday_evening = False
    sunday_morning = False
    sunday_afternoon = False
    # sunday_evening = False

    # program
    if saturday_morning:
        # Saturday from 9am to 12pm
        categorie_selected = df_tbl_gymnastes['Categorie'].isin(['Niveau 3 U13',
                                                                'Niveau 3 13+', 'Élite 3'])
    elif saturday_afternoon:
        # Saturday from 1pm to 5pm
        categorie_selected = df_tbl_gymnastes['Categorie'].isin(['Niveau 5',
                                                                'National Ouvert', 'Junior', 'Senior'])
    elif saturday_evening:
        # Saturday from 5h45pm to 8pm
        categorie_selected = df_tbl_gymnastes['Categorie'].isin(['Niveau 4 U13',
                                                                'Niveau 4 13+'])
    elif sunday_morning:
        # Sunday 8am to 12pm
        categorie_selected = df_tbl_gymnastes['Categorie'].isin(['Niveau 2A',
                                                                'Niveau 2C'])
    elif sunday_afternoon:
        # Sunday 12.30pm to 3.45pm
        categorie_selected = df_tbl_gymnastes['Categorie'].isin(['Niveau 2B',
                                                                'Niveau 2D'])
    return categorie_selected


while True:
    local_database = copy_mdb_file(ORIGINAL_DATABASE_PATH, PATH_DESTINATION)
    df_tbl_notes_with_gymnastes = get_info_to_be_displayed_from_database(local_database)
    send_to_google_spreadsheet_via_api(PRIVATE_KEY_JSON, df_tbl_notes_with_gymnastes)
    print('Website updated, waiting before next iteration... last update at ',
          time.strftime('%Hh%m', time.localtime()))
    time.sleep(30)