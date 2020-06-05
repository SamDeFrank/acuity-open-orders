import os
import data
import web
import utils

settings = utils.load_user_settings()

USERNAME        = settings['username']
PASSWORD        = settings['password']
TSV_PATH        = settings['download_path']
SAVE_PATH       = settings['save_path']
EXCEL_PATH      = settings['excel_path']

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
# print("Parsing data into Excel file")
# excel_file_name = data.main(TSV_PATH, SAVE_PATH)
# print("Done!")
# print("")
# print("Opening new file with Excel")

#check for most recent file here


#Remove file from downloads folder
os.remove(TSV_PATH)

#Open file in excel
os.system(r'start "{}" "{}"'.format(EXCEL_PATH, SAVE_PATH + excel_file_name))
