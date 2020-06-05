import os

dirname = os.path.dirname(__file__)
filename = os.path.join(dirname, 'settings.txt')

def load_user_settings():
  """Load user settings from .txt file into a python dict"""
  with open(filename, 'r') as file:
    settings = dict([x.split(":", 1) for x in file.read().strip().split("\n")])

    for k in settings:
     settings[k] = settings[k].strip()

  return settings

def name_this_later():
  today = datetime.date.today()
  last_monday = today - datetime.timedelta(days=today.weekday())

  delta = today - last_monday

  date_list = [base - datetime.timedelta(days=x) for x in range(delta.days)]

  print(date_list)