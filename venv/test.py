"""from re import search
a_string = "this is a ki_te message"
result = search(r"\(([A-Za-z0-9_]+)\)", a_string)
print(result)
print(result.group(0))
print(result.group(1))

(search(r"\(([A-Za-z0-9_]+)\)", wss.cell(row = i, column = 8).value)).group(0)
wss.cell(row = i, column = 8).value


wss1.cell(row = i, column = 6).value


(search(r"\(([A-Za-z0-9_]+)\)", wss1.cell(row = i, column = 6).value)).group(0)


with open("test.txt", encoding = 'utf-8') as f:
    data = f.readlines()

from json import loads

with open("employee.txt",'r',encoding = 'utf-8') as f:
   empdump = f.readline()

employee = loads(empdump)
print(employee)

start = input("Start date: ")
end = input("End date: ")
month = input("Month: ")
year = input("Year: ")
my = "-" + month + "-" + year

daterange = [str(x)+my for x in range(int(start),int(end)+1)]
print (range)

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

wb = Workbook()

dest_filename = 'empty_book.xlsx'

ws1 = wb.active
ws1.title = "range names"

for row in range(1, 40):
    ws1.append(range(600))

ws2 = wb.create_sheet(title="Pi")

ws2['F5'] = 3.14

ws3 = wb.create_sheet(title="Data")
for row in range(10, 20):
    for col in range(27, 54):
        _ = ws3.cell(column=col, row=row, value="{0}".format(get_column_letter(col)))
print(ws3['AA10'].value)

wb.save(filename = dest_filename)


fnamecase = input("Name of case file: ")
fnamechg1 = input("Name of 1st change file: ")
fnamechg2 = input("Name of 2nd change file: ")
fullpath = "C:\\Users\\utsakuma\\Desktop\\scrypt\\"
fnames = [fullpath + name for name in [fnamecase, fnamechg1, fnamechg2]]
print(fnames)
"""
"""
import datetime

daten = datetime.date(2021, 2, 28)
date = datetime.date.toordinal(daten)
date1 = date + 1
date2 = datetime.date.fromordinal(date1)
print(daten, date, date1, date2)


# importing pandas as pd
import pandas as pd

# generating all dates in given range
# with increment by days
allDates = pd.date_range('2020-06-27', '2020-07-03')

# converting dates to series
series = allDates.to_series()

print(series.dt.isocalendar().week)
#print(type(series.dt.isocalendar().week))
"""

"""
from datetime import datetime
timestamp = datetime.fromordinal(3124600)
print(timestamp)

#timestamp = datetime.toordinal(2021-05-27)
print(timestamp)

timestamp = datetime.fromordinal(3124500)
print(type(timestamp.date()))

timestamp = datetime.toordinal(timestamp.date())
print(type(timestamp))

timestamp = datetime.fromordinal(timestamp)
print(timestamp)

datetime_obj = datetime.strptime("16062021","%d%m%Y")
print(datetime_obj.date())

timestamp = datetime_obj.toordinal()
print(timestamp)

timestamp = timestamp + 1

print(timestamp)
print(datetime.fromordinal(timestamp).date())
start = input("Start date in format(DDMMYYYY): ")
print(type(start))

start = input("Start date in format(DDMMYYYY): ")
datetime_obj = datetime.strptime(start, "%d%m%Y")
datetime_obj = datetime_obj.toordinal()
derange = [datetime.fromordinal(datetime_obj + dates).date().strftime("%d-%m-%Y") for dates in range(0, 7, 1)]
print(derange)
"""
