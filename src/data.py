import os
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.colors import Color

TODAY = datetime.date.today()
OUTPUT_FILENAME = "\\%s Open Orders.xlsx" % datetime.date.today()
SHIP_TOs        = ['Fishers', 'Crawfordsville', 'Des Plaines', 'MPF', 'GPF', 'SEAC']
COLUMN_NAMES    = ["Item Number", "PO Number", "Quantity Ordered", "Quantity Received", "Balance Due", "Need-By Date"]

style = {
  'fonts': {
    'header': Font(size=24, bold=True)
    },
  'borders': {
    'mid': Border(left=Side(style='thin'), right=Side(style='thin'), bottom=Side(style='thin')),
    'underline': Border(bottom=Side(style='thin'))
  },
  'new': {
    'fill' : PatternFill(patternType="solid", fgColor=Color(rgb="FFFFFF00", tint=0.0, type='rgb'),bgColor=Color(rgb="00000000", tint=0.0, type='rgb')),
    'font' : Font()
  },
  'closed': {
    'fill' : PatternFill(),
    'font' : Font(strike=True) 
  },
  'recent': {
    'fill': PatternFill(patternType="solid", fgColor=Color(theme=0,tint=-0.3499862666707358, type='theme'),bgColor=Color(indexed=64,tint=0.0,type='indexed')),
    'font': Font()
  },
  'open': {
    'fill' : PatternFill(),
    'font' : Font()
  }
}


def valid_date(po):
  '''Used as 'key' when sorting purchase orders by due date'''    
  try:
    output = datetime.datetime.strptime(po['Need-By Date'].split()[0], '%d-%b-%Y')
  except ValueError:
    # if the acuity buyer forgets to put a due date on their PO, sort it as if its date were january 1st of 2000. but leave the cell blank
    output = datetime.datetime.strptime("01-jan-2000", '%d-%b-%Y')

  return output

def status(fill, font):
  """assignes status to row based on the current style of the cell"""
  if font.strike:
    return "closed"
  
  #new orders become recent and recent orders stay recent
  elif fill.fgColor.rgb == "FFFFFF00" or fill.fgColor.tint == -0.249977111117893:
    return "recent"
  
  else:
    return "open"

def compare(existing_orders, new_orders):
  updated_orders = existing_orders['orders']
  locations_to_sort = []

  #check for orders to close:
  for ship_to in existing_orders['orders']:
    for po in existing_orders['orders'][ship_to]:
      if po['info']['PO Number'] not in new_orders['po_numbers']:
        po['status'] = 'closed'
  
  #check for new orders to add
  for ship_to in new_orders['orders']:
    for po in new_orders['orders'][ship_to]:
      if po['info']['PO Number'] not in existing_orders['po_numbers']:
        if ship_to not in locations_to_sort:
          locations_to_sort.append(ship_to)
        updated_orders[ship_to].append(po)


  for loc in locations_to_sort:
    updated_orders[loc].sort(key=lambda x : datetime.datetime.strptime(x['info']['Need-By Date'].split()[0], '%m-%d-%Y'))
  
  return updated_orders

def load_tsv(path):
  data_list = []
  orders = {i: [] for i in SHIP_TOs}
  po_numbers = {}
  
  with open(path, 'r') as file:
    raw_data_list = [row.split("\t") for row in file.read().split("\t\n")] 
  
  for i in raw_data_list:
    data_list.append([k.replace('"', "") for k in i])
  
  column_headers = data_list[0]

  #parse data into list of dicts with column headers as the dict keys, then sort it by due date
  order_list = [dict(zip(column_headers, row)) for row in data_list[1:] if len(row) == len(column_headers)]
  order_list = sorted(order_list, key = valid_date)

  #group orders by ship-to-location in a dictionary where ship-to locations are the keys
  for po in order_list:
    
    #if the date string is not empty, reformat it so it's easier to read on the excel sheet   
    if len(po['Need-By Date']) > 0:
      po['Need-By Date'] = datetime.datetime.strptime(po['Need-By Date'].split(" ")[0], '%d-%b-%Y').date().strftime('%m-%d-%Y')

    # po['Balance Due'] = int(po['Quantity Ordered']) - int(po['Quantity Received'])
    po['Balance Due'] = None

    #resolve ship-to location based on whats in the data. eg: both 'P1-CRAWFORDSVILLE' and 'P2-CRAWFORDSVILLE' are grouped under 'Crawfordsville'
    location = [loc for loc in SHIP_TOs if loc.upper() in po['Ship-To Location'].upper()][0]
    
    #trim unnecessary data from each order
    order = {k:v for (k,v) in po.items() if k in COLUMN_NAMES}

    orders[location].append({'info': order, 'status': 'new'})
    po_numbers[order["PO Number"]] = True

  return {'orders': orders, 'po_numbers': po_numbers}

def load_xlsx(path):
  wb = load_workbook(filename = path)
  ws = wb.active
  location = ""
  orders = orders = {i: [] for i in SHIP_TOs}
  po_numbers = {}
  
  for row in ws:

    if row[0].value == None:
      continue

    if row[0].value in SHIP_TOs:
      location = row[0].value
      continue

    values = [cell.value for cell in row]
    order = dict(zip(COLUMN_NAMES, values))

    orders[location].append({'info': order, 'status': status(row[0].fill, row[0].font)})
    po_numbers[order["PO Number"]] = True
  
  return {'orders': orders, 'po_numbers': po_numbers}

def write(wb, orders, created_on, update, settings):
  ws = wb.active
  row_offset = 0
  locations = sorted(orders.keys())

  # set column widths
  ws.column_dimensions["A"].width = 15.14
  ws.column_dimensions["B"].width = 10.00
  ws.column_dimensions["C"].width = 06.75
  ws.column_dimensions["D"].width = 07.71
  ws.column_dimensions["E"].width = 08.43
  ws.column_dimensions["F"].width = 12.29

  ws.merge_cells('G1:I1')
  ws["I1"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

  #the big loop
  for location in locations:
    sz = len(orders[location])

    if(sz > 0):
      row = 1 + row_offset
      ws.row_dimensions[row].height = 30
      ws['A%s' % row].font = style['fonts']['header']
      ws['A%s' % row] = location

      # insert 'created on' date
      if row == 1:
        date_string = "Created On:   {}".format(created_on.strftime('%m-%d-%Y'))
        if update:
          date_string += "\nUpdated On:    {}".format(datetime.date.today().strftime('%m-%d-%Y'))
        ws["I1"] = date_string

      row_offset += 1

      for excelRow in range(1, sz+1):

        # add gridlines to the right of the page for handwritten notes on print out
        if settings['gridlines']:
          ws.cell(row=excelRow + row_offset, column=7).border = style['borders']['underline']
          ws.cell(row=excelRow + row_offset, column=8).border = style['borders']['mid']
          ws.cell(row=excelRow + row_offset, column=9).border = style['borders']['underline']

        #begin writing the business data to sheet   
        for excelCol in range(1, len(COLUMN_NAMES) + 1):

          col_name = COLUMN_NAMES[excelCol-1]
          cell = ws.cell(row=excelRow + row_offset, column=excelCol)
          value = orders[location][excelRow-1]['info'][col_name]
          status = orders[location][excelRow-1]['status']
          
          # convert q ordered and q received values to integers
          if col_name in ["Quantity Ordered", "Quantity Received"]:             
            value = int(value)

          # right justify date
          if col_name == "Need-By Date":
            cell.alignment = Alignment(horizontal="right")

          # generate 'balance due' formula
          if col_name == "Balance Due":
            value = "=if(C{row}-D{row}<0, 0, C{row}-D{row})".format(row=excelRow + row_offset)

          # add horizontal gridlines
          if settings['gridlines']:
            cell.border = style['borders']['underline']


          cell.font = style[status]['font']
          cell.fill = style[status]['fill']
          
          cell.value = value

      row_offset += sz + 1

def create(tsv_path, save_path, settings):

  orders = load_tsv(tsv_path)

  wb = Workbook()
  write(wb, orders['orders'], TODAY, False, settings)

  wb.save('{}{}'.format(save_path, OUTPUT_FILENAME))

  return OUTPUT_FILENAME

def update(tsv_path, xlsx_path, settings):

  created_on = datetime.datetime.strptime(xlsx_path.split("\\")[-1].split(" ")[0], "%Y-%m-%d").date()

  new_orders = load_tsv(tsv_path)
  existing_orders = load_xlsx(xlsx_path)

  os.remove(xlsx_path)

  updated_orders = compare(existing_orders, new_orders)

  wb = Workbook()
  write(wb, updated_orders, created_on, True, settings)

  wb.save('{}'.format(xlsx_path))

  return xlsx_path
