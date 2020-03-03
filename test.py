import csv
import pyodbc
import os.path
import pandas
import numpy

# set up some constants
MDB = os.path.join('C:\\', 'Users', 'admin', 'Documents', 'CoupeQuebec', '2020', 'GAM-CPS_2019-2020.mdb')
DRV = 'Microsoft Access Driver (*.mdb, *.accdb)'
PWD = 'pw'

# connect to db
# connect to db
con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, MDB, PWD))
cur = con.cursor()
# con = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\admin\Documents\CoupeQuebec\2020\GAM-CPS_2019-2020.mdb')
# format(DRV, MDB, PWD))


# run a query and get the results
SQL = 'SELECT * FROM tblGymnaste;'  # your query goes here
rows = numpy.asarray(cur.execute(SQL).fetchall())
cur.close()
con.close()

df = pandas.DataFrame(rows, columns=['idGymnaste', 'NoAffiliation', 'Nom', 'Prenom', 'NomClub', 'Categorie',
                                     'Prov', 'Age', 'Equipe', 'Cat_Equipe'])

df.to_csv(r'export_dataframe.csv')  # Don't forget to add '.csv' at the end of the path

print(df)

# you could change the mode from 'w' to 'a' (append) for any subsequent queries
#with open('mytable.csv', 'wb') as fou:
#    csv_writer = csv.writer(fou)  # default field-delimiter is ","
#    csv_writer.writerows(rows)
