import os
import datetime

# path  = "H:\\Sam\\Documents\\acuity_open_orders"
# index = 0

# files = [f for r, d, f in os.walk(path)]

# print(files)


 
# today = datetime.date.today()
# last_monday = today - datetime.timedelta(days=today.weekday())
# coming_monday = today + datetime.timedelta(days=-today.weekday(), weeks=1)
# print("Today:", today)
# print("Last Monday:", last_monday)
# print("Coming Monday:", coming_monday)


def name_this_later(save_path):
  today = datetime.date.today()
  last_monday = today - datetime.timedelta(days=today.weekday())

  delta = today - last_monday

  date_list = [today - datetime.timedelta(days=x) for x in range(delta.days + 1)]
  filename_list = [f"\\{x.strftime('%Y-%m-%d')} Open Orders.xlsx" for x in date_list]

  for k in filename_list:
    path = save_path + k
    if os.path.isfile(path):
      print(path, "found")
    else:
      print(path, "not found")
  
  # return False

name_this_later(r"H:\Sam\Documents\acuity_open_orders")


