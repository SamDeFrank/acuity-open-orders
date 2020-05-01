def load_user_settings():
  """Load user settings from .txt file into a python dict"""
  with open('./settings.txt', 'r') as file:
    settings = dict([x.split(":", 1) for x in file.read().strip().split("\n")])

  for k in settings:
    settings[k] = settings[k].strip()
  print("loading settings")
  return settings

settings = load_user_settings()
