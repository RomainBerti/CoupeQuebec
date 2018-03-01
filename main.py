# This program is to populate online website with live results during the competition
# main program to fetch marks from MS-access database and push to website

import time
import os
from shutil import copyfile
from pathlib import Path


#def main():

### 1/ Copy the sync database file to local Drive every hour

pathToOriginalDatabase = '/Users/rpgb/Dropbox/CoupeQuebec/2018/BDD-live/2eCOupeQcGAM.mdb'
pathDestination = '/Users/rpgb/Documents/CoupeQuebec/2018'
timeStamp = time.strftime("%Y%m%d_%Hh", time.localtime())
fileToCopy = pathDestination + '/2eCOupeQcGAM_' + timeStamp + '.mdb'

if ~os.path.isfile(fileToCopy):
    # file doesn't already exists
    copyfile(pathToOriginalDatabase, fileToCopy)
