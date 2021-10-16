"""# importing openpyxl module
import openpyxl as xl


# opening the source excel file
filename = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\cases.xlsx"
wb1 = xl.load_workbook(filename)
wss = wb1.worksheets[0]

# opening the destination excel file
filename1 = "C:\\Users\\utsakuma\\Documents\\Case\\2ndApril\\actcase.xlsx"
wb2 = xl.load_workbook(filename1)
wsd = wb2.active
"""
from re import search

def actcase(wss,wsd,employees):
    # calculate total number of rows and
    # columns in source excel file
    mr = wss.max_row
    mc = wss.max_column
    incS = ["Pending", "Work In Progress"] #incident status
    
    # to compare data
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
      #print(str(wss.cell(row = i, column = 17).value))
      #print(str(wss.cell(row = i, column = 18).value))
      #print(str(wss.cell(row = i, column = 19).value))
      if (str(wss.cell(row = i, column = 18).value) in incS) and (result.group(1) in employees):
        print(str(wss.cell(row = i, column = 18).value), wss.cell(row = i, column = 8).value)
        for j in range (2, mc + 1):
          # reading cell value from source excel file
          c = wss.cell(row = i, column = j)
          # writing the read value to destination excel file
          wsd.cell(row = k, column = j - 1).value = c.value
    
        k += 1
# saving the destination excel file
#wb2.save(str(filename1))
