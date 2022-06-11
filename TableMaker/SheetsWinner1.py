import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("GoogleSheetsAPI.json", scope)
client = gspread.authorize(creds)

spreadsheet2 = client.open("RCVSpreadsheet").worksheet("President Trial 1 Table")

limit = 1
while spreadsheet2.cell(limit, 7).value is not None:
    limit = limit + 1
print(limit)
time.sleep(1)

i = 2

while i < limit - 1:
    if int(float(spreadsheet2.cell(i, 7).value)) < 50:
        i = i + 1
        time.sleep(1)
    if int(float(spreadsheet2.cell(i, 7).value)) > 50:
        print(str(spreadsheet2.cell(i, 1).value) + " is the winner")
        time.sleep(1)
        break
print(i)
print(limit)
if (i+1) >= limit:
    print("No winner found. Initiate next table")

column_row = "D1"

# spreadsheet2.update_cell(9, 1, '=MID(Sheet1!' + column_row + ',FIND("[",Sheet1!' + column_row + ')+1,FIND("]",Sheet1!'+ column_row + ')-FIND("[",Sheet1!' + column_row + ')-1)')