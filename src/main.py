from config import config
from settings import settings
import os
import data
print(settings)

TSV_PATH        = settings['dev_path'] if len(settings['dev_path']) > 0 else settings['download_path']
SAVE_PATH       = settings['save_path'] if len(settings['save_path']) > 0 else ""
EXCEL_COLUMNS   = config['include_columns']
EXCEL_PATH      = settings['excel_path']
INCLUDE_BALANCE = settings['include_computed_balance']


excel_file_name = data.main(TSV_PATH, EXCEL_COLUMNS, INCLUDE_BALANCE, SAVE_PATH)

os.system(r'start "{}" "{}"'.format(EXCEL_PATH, SAVE_PATH + excel_file_name))
