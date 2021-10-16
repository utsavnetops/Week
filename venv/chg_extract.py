"""# importing openpyxl module
import openpyxl as xl
from re import search

# opening the source excel file
filename = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\1.xlsx"
wb1 = xl.load_workbook(filename)
wss1 = wb1.worksheets[0]

filename = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\2.xlsx"
wb2 = xl.load_workbook(filename)
wss2 = wb1.worksheets[0]

# opening the destination excel file
filename1 = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\chg.xlsx"
wb3 = xl.load_workbook(filename1)
wsd = wb3.active
"""

from re import search


def chgext(wss1,wss2,wsd,date,employees):
    # calculate total number of rows and
    # columns in source excel file
    mr1 = wss1.max_row
    mc1 = wss1.max_column
    mr2 = wss2.max_row
    mc2 = wss2.max_column
    # to compare data
    #week = ["2021-04-18", "2021-04-19", "2021-04-20", "2021-04-21", "2021-04-22", "2021-04-23", "2021-04-24"]
    #employee = ["Himanshu Handa (hihanda)", "Jithu Poovanolickal (jipoovan)", "PHANI SAI MUKESH MARIETTI (pmariett)", "Ram Uppaluri (ruppalur)", "Arun Kumar C (arkumarc)", "Utsav Kumar (utsakuma)", "Sandeep Chandroth Kavalan (schandr3)", "Surjeet Kumar (surjekum)"]
    unique = []
    k = 1
    # copying the cell values from source
    # excel file to destination excel file
    print ("file 1")
    for i in range (1, mr1 + 1):
      #print(wss1.cell(row = i, column = 2).value ,str(wss1.cell(row = i, column = 10).value)[0:10], wss1.cell(row=i, column=6).value)
      val = wss1.cell(row=i, column=6).value
      #print(val)
      if val == None:
          val = "ss(sample)ll"
      result = search(r"\(([A-Za-z0-9_]+)\)", val)

        # svalue = search(r"\(([A-Za-z0-9_]+)\)", val)
      if not (result):
          print('empty found')
          continue

      #print(str(wss1.cell(row=i, column=10).value)[0:10])

      #print(type(result.group(1)),end='\t')
      #print(result.group(1))
      #print(str(wss1.cell(row = i, column = 10).value)[0:10], result.group(1))
      if (str(wss1.cell(row = i, column = 10).value)[0:10] in date) and (result.group(1) in employees):
    
        print("Loop 1", wss1.cell(row = i, column = 2).value, wss1.cell(row = i, column = 11).value, str(wss1.cell(row = i, column = 10).value), wss1.cell(row = i, column = 6).value)
    
        if((wss1.cell(row = i, column = 2).value) in unique):
          print("skip")
          continue
        unique.append(str(wss1.cell(row = i, column = 2).value))
        for j in range (2, mc1 + 1):
          # reading cell value from source excel file
          c = wss1.cell(row = i, column = j)
          # writing the read value to destination excel file
          wsd.cell(row = k, column = j - 1).value = c.value
        k += 1
    
    # copying the cell values from source
    # excel file to destination excel file
    print ("file 2")
    for i in range (1, mr2 + 1):
      #print(wss2.cell(row = i, column = 2).value, str(wss2.cell(row = i, column = 10).value)[0:10], wss2.cell(row=i, column=6).value)
      val = wss2.cell(row=i, column=6).value
      #print(val)
      if val == None:
          val = "ss(sample)ll"
      result = search(r"\(([A-Za-z0-9_]+)\)", val)

        # svalue = search(r"\(([A-Za-z0-9_]+)\)", val)
      if not (result):
          print('empty found')
          continue
      #print(str(wss2.cell(row = i, column = 10).value)[0:10], result.group(1))
      if (str(wss2.cell(row = i, column = 10).value)[0:10] in date) and (result.group(1) in employees):
    
        print("Loop 2", wss2.cell(row = i, column = 2).value, wss2.cell(row = i, column = 11).value, str(wss2.cell(row = i, column = 10).value), wss2.cell(row = i, column = 6).value)
        if((wss2.cell(row = i, column = 2).value) in unique):
          print("skip")
          continue
        unique.append(str(wss2.cell(row = i, column = 2).value))
        for j in range (2, mc2 + 1):
    
          # reading cell value from source excel file
          c = wss2.cell(row = i, column = j)
    
          # writing the read value to destination excel file
          wsd.cell(row = k, column = j - 1).value = c.value
    
        k += 1
# saving the destination excel file
#wb3.save(str(filename1))