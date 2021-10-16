"""# importing openpyxl module
import openpyxl as xl
from re import search

# opening the source excel file
filename = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\cases.xlsx"
wb1 = xl.load_workbook(filename)
wss = wb1.worksheets[0]

# opening the destination excel file
filename1 = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\case.xlsx"
wb2 = xl.load_workbook(filename1)
wsd = wb2.active
"""
from re import search


def closedcases(wss,wsd,date,employees):
    # calculate total number of rows and
    # columns in source excel file
    mr = wss.max_row
    mc = wss.max_column

    # to compare data
    #week = ["2021-04-18 00:00:00", "2021-04-19 00:00:00", "2021-04-20 00:00:00", "2021-04-21 00:00:00", "2021-04-22 00:00:00", "2021-04-23 00:00:00", "2021-04-24 00:00:00"]
    #employee = ["Himanshu Handa (hihanda)", "Jithu Poovanolickal (jipoovan)", "PHANI SAI MUKESH MARIETTI (pmariett)", "Ram Uppaluri (ruppalur)", "Arun Kumar C (arkumarc)", "Utsav Kumar (utsakuma)", "Sandeep Chandroth Kavalan (schandr3)", "Surjeet Kumar (surjekum)"]
    k = 1
    # copying the cell values from source
    # excel file to destination excel file
    for i in range (1, mr + 1):
      val = wss.cell(row = i, column = 8).value
      if val == None:
          val = "ss(sample)ll"
      result = search(r"\(([A-Za-z0-9_]+)\)", val)

      #svalue = search(r"\(([A-Za-z0-9_]+)\)", val)
      if not (result):
          print('empty found')
          continue

      if (str(wss.cell(row = i, column = 12).value)[0:10] in date) and (result.group(1) in employees):
        print(str(wss.cell(row = i, column = 12).value), wss.cell(row = i, column = 8).value)
        for j in range (2, mc + 1):
    
          # reading cell value from source excel file
          c = wss.cell(row = i, column = j)
    
          # writing the read value to destination excel file
          wsd.cell(row = k, column = j - 1).value = c.value
    
        k += 1
# saving the destination excel file
#wb2.save(str(filename1))