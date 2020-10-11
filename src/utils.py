import os
import datetime

dirname = os.path.dirname(__file__)
filename = os.path.join(dirname, 'settings.txt')

def load_user_settings():
  """Load user settings from .txt file into a python dict"""
  with open(filename, 'r') as file:
    settings = dict([x.split(":", 1) for x in file.read().strip().split("\n")])

    for k in settings:
     settings[k] = settings[k].strip()

  return settings

def current_report(save_path):
  """determines if a file was created since the most recent monday. returns the path if it exists, returns 0 if it doesnt"""
  today = datetime.date.today()
  last_monday = today - datetime.timedelta(days=today.weekday())

  delta = today - last_monday

  date_list = [today - datetime.timedelta(days=x) for x in range(delta.days + 1)]
  filename_list = [f"\\{x.strftime('%Y-%m-%d')} Open Orders.xlsx" for x in date_list]

  for k in filename_list:
    path = save_path + k
    if os.path.isfile(path):
      return path
  
  return 0

def get_download_path():
  """Returns the default downloads path for linux or windows"""
  if os.name == 'nt':
    import winreg
    sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
    downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
      location = winreg.QueryValueEx(key, downloads_guid)[0]
    return location
  else:
    return os.path.join(os.path.expanduser('~'), 'downloads')