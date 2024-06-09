import csv
import gspread
import time
import gspread.exceptions

file = f"DMAT - Sheet1.csv"

entries = []
sum=0

def finmgmt(file):
    with open(file, mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)
        for row in csv_reader:
            date = row[0]
            mode = row[1]
            particulars = row[2]
            deposits = row[3]
            withdrawals = row[4]
            balance = row[5]
            # print(sum)
            entry = ((date, mode, particulars,deposits, withdrawals,balance))
            print(entry)
            entries.append(entry)
        return entries

sa = gspread.service_account()
sh = sa.open("pyPt2")

worksheets = sh.worksheets()
print("Available Worksheets:")
for worksheet in worksheets:
    print(worksheet.title)


# wks = sh.worksheet(f"bkt.csv")
rows = finmgmt(file)
# worksheet.insert_row([1,2,3], 13)


for row in rows:
    print(row[4])  # Print the value of withdrawals for debugging
    if row[4] == '0':
        try:
            # Insert the row into the worksheet
            sh.sheet1.insert_row([row[0], row[1], row[2], row[3], row[4], row[5]], 2)
            time.sleep(2)
        except gspread.exceptions.APIError as e:
            print(f"Error inserting row {row} into Google Sheets: {e}")


