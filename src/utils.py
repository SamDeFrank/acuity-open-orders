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

settings = load_user_settings()
