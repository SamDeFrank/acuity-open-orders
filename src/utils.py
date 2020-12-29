from datetime import datetime
import os

dirname = os.path.dirname(__file__)
filename = os.path.join(dirname, 'settings.txt')

def load_user_settings():
  """Load user settings from .txt file into a python dict"""

  with open(filename, 'r') as file:
    settings = dict([x.split(":", 1) for x in file.read().strip().split("\n")])

  settings = { k: v.strip() for k, v in settings.items() }

  return settings


def current_report(path):
  """determines if a file was created since the most recent monday. returns the path if it exists, returns 0 if it doesn't"""
  today = datetime.today()

  def key(x):
    return os.stat(os.path.join(path, x.name)).st_ctime

  files = sorted(os.scandir(path), key=key)

  if len(files) > 0:
    most_recent = datetime.fromtimestamp(os.stat(path + files[-1].name).st_ctime)
  else:
    return 0

  delta = today - most_recent

  if delta.days <= today.weekday():
    return os.path.join(path, files[-1])
  else:
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