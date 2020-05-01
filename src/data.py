import datetime
from utils import settings
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.styles.borders import Border, Side

OUTPUT_FILENAME = "\\%s Open Orders.xlsx" % datetime.date.today()
SHIP_TOs        = ['Fishers', 'Crawfordsville', 'Des Plaines', 'MPF', 'GPF', 'SEAC']
COLUMN_NAMES    = ["Item Number", "PO Number", "Quantity Ordered", "Quantity Received", "Balance Due", "Need-By Date"]

header_style = {
  'size': 24,
  'bold': True
}

borders = {
  'mid': Border(left=Side(style='thin'), right=Side(style='thin'), bottom=Side(style='thin')),
  'underline': Border(bottom=Side(style='thin'))
} 

def valid_date(po):
  '''This function is used as the 'key' when sorting purchase orders by due date'''    
  try:
    output = datetime.datetime.strptime(po['Need-By Date'].split()[0], '%d-%b-%Y')
  except ValueError:
    # if the acuity buyer forgets to put a due date on their PO, sort it as if its date were january 1st of 2000. but leave the cell blank
    output = datetime.datetime.strptime("01-jan-2000", '%d-%b-%Y')

  return output

def main(tsv_path, save_path): 
  
  data_list = []
  po_by_ship_to = {}
  
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


    po['Balance Due'] = int(po['Quantity Ordered']) - int(po['Quantity Received'])

    #resolve ship-to location based on whats in the data. eg: both 'P1-CRAWFORDSVILLE' and 'P2-CRAWFORDSVILLE' are grouped under 'Crawfordsville'
    ship_to = [loc for loc in SHIP_TOs if loc.upper() in po['Ship-To Location'].upper()][0]

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
  locations = sorted(po_by_ship_to.keys())

  # set column widths
  ws.column_dimensions["A"].width = 15.14
  ws.column_dimensions["B"].width = 10.00
  ws.column_dimensions["C"].width = 06.75
  ws.column_dimensions["D"].width = 07.71
  ws.column_dimensions["E"].width = 08.43
  ws.column_dimensions["F"].width = 12.29

  #the big loop
  for location in locations:
    sz = len(po_by_ship_to[location])

    if(sz > 0):
      row = 1 + row_offset
      ws.row_dimensions[row].height = 30
      ws['A%s' % row].font = Font(size=header_style['size'], bold=header_style['bold'])
      ws['A%s' % row] = location

      # insert 'created on' date
      if row == 1:
        ws["I1"] = "Created On:   {}".format(datetime.date.today().strftime('%m-%d-%Y'))
        ws["I1"].alignment = Alignment(horizontal="right", vertical="top")
      
      row_offset += 1

      for excelRow in range(1, sz+1):

        # add gridlines to the right of the page for handwritten notes on print out
        if settings['gridlines']:
          ws.cell(row=excelRow + row_offset, column=7).border = borders['underline']
          ws.cell(row=excelRow + row_offset, column=8).border = borders['mid']
          ws.cell(row=excelRow + row_offset, column=9).border = borders['underline']

        #begin writing the business data to sheet   
        for excelCol in range(1, len(COLUMN_NAMES) + 1):    
          col_name = COLUMN_NAMES[excelCol-1]
          cell = ws.cell(row=excelRow + row_offset, column=excelCol)
          value = po_by_ship_to[location][excelRow-1][col_name]
          
          # convert q ordered and q received values to integers
          if col_name in ["Quantity Ordered", "Quantity Received"]:             
            value = int(value)

          # right justify date
          if col_name == "Need-By Date":
            cell.alignment = Alignment(horizontal="right")

          # generate 'balance due' formula
          if col_name == "Balance Due":
            value = "=C{row} - D{row}".format(row=excelRow + row_offset)

          # add horizontal gridlines
          if settings['gridlines']:
            cell.border = borders['underline']
          
          cell.value = value
          
      row_offset += sz + 1

  wb.save('{}{}'.format(save_path, OUTPUT_FILENAME))
  return OUTPUT_FILENAME

 
if __name__ == "__main__":
  import os
  main(settings['dev_path'], settings['save_path'])
  os.system(r'start "{}" "{}{}"'.format(settings['excel_path'], settings['save_path'], OUTPUT_FILENAME))
