# This program is to populate online website with live results during the competition
# main program to fetch marks from MS-access database and push to website

import pandas
import subprocess
import time
from shutil import copyfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyodbc
import os.path
import numpy

# constants
PATH_DESTINATION = os.path.join('C:\\', 'Users', 'admin', 'Documents', 'CoupeQuebec', '2020')
PRIVATE_KEY_JSON = os.path.join(PATH_DESTINATION, 'MyProject8533_2020.json')
ORIGINAL_DATABASE_PATH = os.path.join(PATH_DESTINATION, 'GAM-CPS_2019-2020.mdb')
# set up some constants
MDB = os.path.join('C:\\', 'Users', 'admin', 'Documents', 'CoupeQuebec', '2020', 'GAM-CPS_2019-2020.mdb')
MS_ACCESS_DRIVER = 'Microsoft Access Driver (*.mdb, *.accdb)'

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
        # connect to db
        con = pyodbc.connect('DRIVER={};DBQ={}'.format(MS_ACCESS_DRIVER, path_to_local_database))
        cur = con.cursor()
        SQL = 'SELECT * FROM ' + table_name + ';'  # SQL query
        rows = numpy.asarray(cur.execute(SQL).fetchall())
        columns = [column[0] for column in cur.description]
        cur.close()
        con.close()
        df = pandas.DataFrame(rows, columns=columns)
        export_name = PATH_DESTINATION + '/exportedTable_' + table_name + '.csv'
        df.to_csv(export_name, index=False)  # Don't forget to add '.csv' at the end of the path

    # load csv in dataframe to do join on data and export result in csv with pandas
    # in the df names, images keep the original table name
    df_tbl_notes = pandas.read_csv(os.path.join(PATH_DESTINATION, 'exportedTable_tblNote.csv'))
    df_tbl_gymnastes = pandas.read_csv(os.path.join(PATH_DESTINATION, 'exportedTable_tblGymnaste.csv'))
    # switch NoAffiliation and idGymnaste which are inverted
    df_tbl_gymnastes.columns = ['NoAffiliation', 'idGymnaste', 'Nom', 'Prenom', 'NomClub', 'Categorie',
                                'Prov', 'Age', 'Equipe', 'Cat_Equipe']

    # Select lines for the 2eme coupe quebec CPS
    # NB: the competition name changes every year, make sure to use the correct one
    df_tbl_notes = df_tbl_notes[df_tbl_notes['Competition'] == '2e Coupe Qc 2019-2020 - CPS']

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
    print(df_tbl_notes_with_gymnastes.shape)
    cell_list = ws.range('A1:I' + str(df_tbl_notes_with_gymnastes.shape[0]))

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

# Make sure the destination page on googledrive is long enough, otherwise you will get a range error

