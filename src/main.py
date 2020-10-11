import os
import data
import web
import utils

settings = utils.load_user_settings()

USERNAME        = settings['username']
PASSWORD        = settings['password']
SAVE_PATH       = settings['save_path']
EXCEL_PATH      = settings['excel_path']
WEBDRIVER_PATH  = settings['webDriver_path']
DOWNLOAD_PATH   = utils.get_download_path()
TSV_PATH        = DOWNLOAD_PATH + "\export.tsv"

#Remove old tsv from downloads folder if it exists
if os.path.isfile(TSV_PATH):
    os.remove(TSV_PATH)

#Fetch .tsv file from Acuity supplier portal
downloading = web.fetchTSV(USERNAME, PASSWORD, DOWNLOAD_PATH, WEBDRIVER_PATH)

#wait for file to download
print("downloading...")
while downloading:
    if os.path.isfile(TSV_PATH):
        downloading = False
        print("Download complete")
    else:
        pass

#determine if there is a recent file to update, or a new file needs to be generated.
file_to_update = utils.current_report(SAVE_PATH)

if file_to_update != 0:
    print("Updating most recent order report.")
    excel_file_name = data.update(TSV_PATH, file_to_update, settings)
    print("Done!")
    print("Opening new file with Excel")
    
    os.system(r'start "{}" "{}"'.format(EXCEL_PATH, excel_file_name))
else:
    print("Parsing data into Excel file")
    excel_file_name = data.create(TSV_PATH, SAVE_PATH, settings)
    print("Done!")
    print("Opening new file with Excel")

    #Open file in excel
    os.system(r'start "{}" "{}"'.format(EXCEL_PATH, SAVE_PATH + excel_file_name))