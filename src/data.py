import datetime
from settings import settings
from config import config
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

OUTPUT_FILENAME = "\\%s Open Orders.xlsx" % datetime.date.today()

def valid_date(po):
  '''This function is used as the 'key' when sorting purchase orders by due date'''    
  try:
    output = datetime.datetime.strptime(po['Need-By Date'].split()[0], '%d-%b-%Y')
  except ValueError:
    output = datetime.datetime.strptime("01-jan-2000", '%d-%b-%Y')

  return output

def main(tsv_path, cols, include_balance, save_path): 
  
  data_list = []
  po_by_ship_to = {}
  cols = cols[:-1] + ['Balance Due'] + cols[-1:] if include_balance else cols
  
  with open(tsv_path, 'r') as file:
    raw_data_list = [row.split("\t") for row in file.read().split("\t\n")] 
  
  for i in raw_data_list:
    data_list.append([k.replace('"', "") for k in i])
  
  column_headers = data_list[0]

  #parse data into list of dicts with column headers as the dict keys, then sort it by due date
  data_dict = [dict(zip(column_headers, row)) for row in data_list[1:] if len(row) == len(column_headers)]
  data_dict = sorted(data_dict, key = valid_date)

  #group orders by ship-to-location in a dictionary where ship-to locations are the keys
  for po in data_dict:

    #if the date string is not empty, reformat it so it's easier to read on the excel sheet   
    if len(po['Need-By Date']) > 0:
      po['Need-By Date'] = datetime.datetime.strptime(po['Need-By Date'].split(" ")[0], '%d-%b-%Y').date().strftime('%m-%d-%Y')

    if include_balance:
      po['Balance Due'] = int(po['Quantity Ordered']) - int(po['Quantity Received'])

    #resolve ship-to location based on whats in the data. eg: both 'P1-CRAWFORDSVILLE' and 'P2-CRAWFORDSVILLE' are grouped under 'Crawfordsville'
    ship_to = [loc for loc in config['ship_to_locations'] if loc.upper() in po['Ship-To Location'].upper()][0]

    if ship_to in po_by_ship_to.keys():
      po_by_ship_to[ship_to].append(po)
    else:
      po_by_ship_to[ship_to] = [po]
  #------------------------------------------------------------
  #----------------------BEGIN EXCEL WORK----------------------
  #------------------------------------------------------------

  wb = Workbook()
  ws = wb.active
  row_offset = 0
  locations = settings['custom_location_order'] if len(settings['custom_location_order']) > 0 else sorted(po_by_ship_to.keys())

  header = config['styles']['header']

  for location in locations:
    sz = len(po_by_ship_to[location])

    if(sz > 0):
      row = 1 + row_offset
      ws.row_dimensions[row].height = 30
      ws['A%s' % row].font = Font(size=header['size'], bold=header['bold'])
      ws['A%s' % row] = location

      if row == 1:
        ws["I1"] = "Created On:   {}".format(datetime.date.today().strftime('%m-%d-%Y'))
        ws["I1"].alignment = Alignment(horizontal="right", vertical="top")
      
      row_offset += 1

      for excelRow in range(1, sz+1):
        for excelCol in range(1, len(cols) + 1):
          cell_value = po_by_ship_to[location][excelRow-1][cols[excelCol-1]]
          if cols[excelCol-1] == "Quantity Ordered" or cols[excelCol-1] == "Quantity Received":
            cell_value = int(cell_value)
          ws.cell(row=excelRow + row_offset, column=excelCol, value=cell_value)
      
      row_offset += sz + 1
  wb.save('{}{}'.format(save_path, OUTPUT_FILENAME))
  return OUTPUT_FILENAME

 
if __name__ == "__main__":
  import os
  main(settings['dev_path'], config['include_columns'], True, settings['save_path'])
  os.system(r'start "{}" "{}{}"'.format(settings['excel_path'], settings['save_path'], OUTPUT_FILENAME))
