import os
import json
import win32com.client as win32
from openpyxl import Workbook

f = open('train.json')
jsonData = json.load(f)
rows = []
ctr = 0
##Analysis the json file by name and tweet then append it to the list rows.
for record in jsonData:
    name = record["profile"]["screen_name"]
    tweet = record["tweet"]
    if tweet is not None:
        for t in tweet:
            if (ctr == 5000):
                break
            if (t.startswith("RT")):
                t = t.split(":")
                t = t[0]
                rows.append([name, t[4:]])
                ctr += 1


ExcelApp = win32.Dispatch('Excel.Application')
ExcelApp.visible = True

wb = ExcelApp.Workbooks.Add()
ws = wb.Worksheets(1)

header_labels = ('name', 'tweet')

# insert header labels
for index, val in enumerate(header_labels):
    ws.Cells(1, index + 1).Value = val

##insert Records to excel file
row_tracker = 2
column_size = len(header_labels)

for row in rows:
    ws.Range(
        ws.Cells(row_tracker, 1),
        ws.Cells(row_tracker, column_size)
    ).value = row
    row_tracker += 1

