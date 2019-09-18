from config import config
from settings import settings
import os
import data
import web

USERNAME        = settings['username']
PASSWORD        = settings['password']
TSV_PATH        = settings['dev_path'] if len(settings['dev_path']) > 0 else settings['download_path']
SAVE_PATH       = settings['save_path'] if len(settings['save_path']) > 0 else ""
EXCEL_COLUMNS   = config['include_columns']
EXCEL_PATH      = settings['excel_path']
INCLUDE_BALANCE = settings['include_computed_balance']

#Fetch .tsv file from Acuity supplier portal
downloading = web.fetchTSV(USERNAME, PASSWORD)

#wait for file to download
print("Waiting for file to finish downloading")
while downloading:
    if os.path.isfile(TSV_PATH):
        downloading = False
        print("Download complete")
    else:
        pass

#Parse data from tsv file into excel format
print("Parsing data into Excel file")
excel_file_name = data.main(TSV_PATH, EXCEL_COLUMNS, INCLUDE_BALANCE, SAVE_PATH)
print("Done!")
print("")
print("Opening new file with Excel")

#Remove file from downloads folder
os.remove(TSV_PATH)

#Open file in excel
os.system(r'start "{}" "{}"'.format(EXCEL_PATH, SAVE_PATH + excel_file_name))
