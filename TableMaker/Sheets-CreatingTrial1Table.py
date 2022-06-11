import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import string

position1 = "President"

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("GoogleSheetsAPI.json", scope)
client = gspread.authorize(creds)

spreadsheets = client.open("RCVSpreadsheet")
survey_results = client.open("RCVSpreadsheet").worksheet("Sheet1")


def phase2_table_maker():
    pass


def initial_table_maker(position1):
    time.sleep(3)
    candidate_number = 2
    candidates = 0
    loop = 1
    while loop == 1:
        if survey_results.cell(1, candidate_number).value is None:
            loop = 2
        time.sleep(1)
        if position1 in str(survey_results.cell(1, candidate_number).value):
            candidate_number = candidate_number + 1
            candidates += 1
        time.sleep(1)
        if position1 not in str(survey_results.cell(1, candidate_number).value):
            candidate_number = candidate_number + 1

    trial_1_table = spreadsheets.add_worksheet(f"{position1} Trial 1 Table", candidates + 2, 10)
    trial_1_table.update_cell(1, 1, position1 + " Candidates")
    trial_1_table.update_cell(1, 7, "First Preference %")
    trial_1_table.update_cell(1, 9, "Total Number of Votes")
    # limits number of responses to 500 (may change if needed be)
    trial_1_table.update_cell(2, 9, '=COUNTA(Sheet1!A2:Sheet1!A500)')
    time.sleep(5)
    # labels each of the row with candidates
    loop = 1
    candidate_number2 = 2
    table1_row = 2
    d = dict(enumerate(string.ascii_uppercase, 1))
    while loop == 1:
        time.sleep(7)
        if position1 not in str(survey_results.cell(1, candidate_number2).value):
            candidate_number2 += 1
            time.sleep(3)
        if position1 in str(survey_results.cell(1, candidate_number2).value):
            row = str(survey_results.cell(1, candidate_number2).row)
            column = int(survey_results.cell(1, candidate_number2).col)
            letter_column = d[column]
            column_row = letter_column + row
            trial_1_table.update_cell(table1_row, 1,
                                      f'=MID(Sheet1!' + column_row + ',FIND("[",Sheet1!' + column_row + ')+1,FIND("]",Sheet1!' + column_row + ')-FIND("[",Sheet1!' + column_row + ')-1)')
            time.sleep(3)
            index_column = 2
            if 1 == 1:
                trial_1_table.update_cell(table1_row, index_column,
                                          f'=countif(Sheet1!{letter_column}:{letter_column}, "First_Choice")')
                index_column += 1
                trial_1_table.update_cell(table1_row, index_column,
                                          f'=countif(Sheet1!{letter_column}:{letter_column}, "Second_Choice")')
                index_column += 1
                trial_1_table.update_cell(table1_row, index_column,
                                          f'=countif(Sheet1!{letter_column}:{letter_column}, "Third_Choice")')
                index_column += 1
                trial_1_table.update_cell(table1_row, index_column,
                                          f'=countif(Sheet1!{letter_column}:{letter_column}, "Fourth_Choice")')
                index_column += 1
                trial_1_table.update_cell(table1_row, index_column,
                                          f'=countif(Sheet1!{letter_column}:{letter_column}, "Fifth_Choice")')
                index_column += 1
                trial_1_table.update_cell(table1_row, index_column, f'=round((B{table1_row}/I2)*100, 2)')
            candidate_number2 += 1
            table1_row += 1
        if survey_results.cell(1, candidate_number2).value is None:
            loop = 2
        time.sleep(1)

    # labels columns with Preference Votes (ex. First Preference Votes, Second Preference Votes
    trial_1_table.update_cell(1, 2, "First Preference")
    time.sleep(3)
    trial_1_table.update_cell(1, 3, "Second Preference")
    trial_1_table.update_cell(1, 4, "Third Preference")
    trial_1_table.update_cell(1, 5, "Fourth Preference")
    time.sleep(3)
    trial_1_table.update_cell(1, 6, "Fifth Preference")

    time.sleep(6)
    # spreadsheet_checker = client.open("RCVSpreadsheet").worksheet(f"{position1} Trial 1 Table")

    limit = 1
    while trial_1_table.cell(limit, 7).value is not None:
        limit = limit + 1

    time.sleep(1)

    i = 2

    while i < limit - 1:
        if int(float(trial_1_table.cell(i, 7).value)) < 50:
            i = i + 1
            time.sleep(2)
        time.sleep(2)
        if int(float(trial_1_table.cell(i, 7).value)) > 50:
            print(str(trial_1_table.cell(i, 1).value) + " is the winner")
            time.sleep(1)
            break
    if (i + 1) >= limit:
        print("No winner found. Creating next table")
        phase2_table_maker()


initial_table_maker("President")
print("Table Created")
