# will most likely end up importing settings to this class

class Common(object):

  ship_tos = ['Fishers', 'Crawfordsville', 'Des Plaines', 'MPF', 'GPF', 'SEAC']


  #still missing tsv-name to excel-name mappings for GPF and SEAC
  ship_to_mapping = {
    "GPF": "GPF",
    "SEAC": "SEAC",
    "MPF" : "MPF",
    "Crawfordsville": "Crawfordsville",
    "Fishers" : "Fishers",
    "Des Plaines" : "Des Plaines",
    "MPF GRUPO (DIR)" : "MPF",
    "P2-CRAWFORDSVILLE IN" : "Crawfordsville",
    "P1-CRAWFORDSVILLE IN" : "Crawfordsville",
    "FISHERS, IN" : "Fishers",
    "DES PLAINES, IL": "Des Plaines",
  }

  col_names = [
    "Item Number",
    "PO Number",
    "Quantity Ordered",
    "Quantity Received",
    "Balance Due",
    "Need-By Date", 
    "G",
    "H",
    "I"
    ]