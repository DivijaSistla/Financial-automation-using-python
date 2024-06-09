from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import BatchHttpRequest
import csv
import gspread
import time
from collections import defaultdict


file = "bankStmt - Sheet1.csv"

entries = []
monTot = defaultdict(float)
CategTot = defaultdict(float)
biggestTs = defaultdict(lambda: {'amount': 0, 'transaction': None})
totAmtCD = 0
totAmtDD = 0

def classify_transaction(description):
    categories = {
        'entertainment': ['netflix subscription', 'zomato order'],
        'food': ['grocery store', 'restaurant', 'zomato order'],
        'income': ['salary deposit', 'refund', 'transfer from friend', 'investment dividend', 'reimbursement', 'consulting fee', 'gift from family'],
        'utilities': ['utility bill payment'],
        'shopping': ['online shopping', 'electronics store', 'online subscription', 'bookstore'],
        'other': ['atm withdrawal', 'transportation', 'upi transfer']
    }

    description_lower = description.lower()

    for category, keywords in categories.items():
        if any(keyword in description_lower for keyword in keywords):
            return category

    return 'unknown'

def parse_date(date):
    try:
        # Assuming the date format is already '%Y-%m-%d'
        return time.strptime(date, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"Invalid date format: {date}")

def get_worksheet(sh, sheet_title):
    worksheets = sh.worksheets()
    for worksheet in worksheets:
        if worksheet.title == sheet_title:
            return worksheet
    return None

def finmgmt(file, sheet_title):
    global totAmtCD, totAmtDD
    totAmtCD = 0
    totAmtDD = 0

    sum = 0
    max_category = None

    sa = gspread.service_account()
    sh = sa.open("pyPt1")

    # Get the worksheet or create a new one if it doesn't exist
    worksheet = get_worksheet(sh, sheet_title)
    if worksheet is None:
        worksheet = sh.add_worksheet(sheet_title, 100, 5)  # Adjust the number of rows and columns as needed

    # Get the current number of rows in the worksheet
    num_rows = len(worksheet.get_all_values())

    # Calculate the starting row for insertion
    start_row = num_rows + 1

    with open(file, mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)

        for row in csv_reader:
            date, transaction_type, amount, description = row
            amt = float(amount)
            sum += amt

            category = classify_transaction(description)

            # Update category total
            CategTot[category] += amt

            # Check if the current category has the maximum sum
            if CategTot[category] > CategTot.get(max_category, 0):
                max_category = category

            # Extract year and month from the date
            date_obj = parse_date(date)
            year, month = date_obj.tm_year, date_obj.tm_mon

            # Update monthly total
            monTot[(year, month)] += amt

            # Track the biggest transaction in each month
            if amt > biggestTs[(year, month)]['amount']:
                biggestTs[(year, month)]['amount'] = amt
                biggestTs[(year, month)]['transaction'] = (date, transaction_type, amount, description, category)

            # Update net amount credited and debited
            if transaction_type.lower() == 'credit':
                totAmtCD += amt
            elif transaction_type.lower() == 'debit':
                totAmtDD += amt

            entry = ((date, amount, transaction_type, description, category))
            print(entry)
            entries.append(entry)

    return entries, max_category, start_row, worksheet


def print_monTot():
    for (year, month), total in monTot.items():
        print(f"Total money spent in {year}-{month:02d}: ${total:.2f}")

def insert_monTot_in(start_row, worksheet):
    for (year, month), total in monTot.items():
        # Insert a new row with month and total expenses
        worksheet.insert_row([f"Total for {year}-{month:02d}", " ", total," "], start_row)
        time.sleep(2)

def insert_max_category_in(max_category, start_row, worksheet):
    # Insert the category with the maximum sum
    worksheet.insert_row(["Category with Max Expense", max_category], start_row+7)
    time.sleep(2)

def insert_biggestTs_in(start_row, worksheet):
    # Insert the header for the biggest transactions section
    worksheet.insert_row(["Biggest Transaction in Each Month", "", "", "", ""], start_row + 8)

    # Insert the details of the biggest transaction in each month
    for (year, month), data in biggestTs.items():
        if data['transaction']:
            date, transaction_type, amount, description, category = data['transaction']
            worksheet.insert_row([f"{year}-{month:02d}", transaction_type, amount, description, category], start_row + 3)
            time.sleep(2)

def insert_net_amounts_in(start_row, worksheet):
    # Insert the header for net amounts section
    worksheet.insert_row(["Net Amounts", "", "", "", ""], start_row + 14)

    # Insert net amount credited and debited
    worksheet.insert_row(["Net Amount Credited", "", "", "", totAmtCD], start_row + 15)
    worksheet.insert_row(["Net Amount Debited", "", "", "", totAmtDD], start_row + 16)

def create_chart_request(sheet_id, chart_title, chart_type, data_range, anchor_cell):
    chart_spec = {
        "title": chart_title,
        "basicChart": {
            "chartType": chart_type,
            "legendPosition": "BOTTOM_LEGEND",
            "axis": [
                {"position": "BOTTOM_AXIS", "format": {"title": "Category"}},
                {"position": "LEFT_AXIS", "format": {"title": "Amount"}}
            ],
            "domains": [{"domain": {"sourceRange": {"sources": [data_range]}}}],
            "series": [{"series": {"sourceRange": {"sources": [data_range]}}}]
        }
    }

    chart_request = {
        "addChart": {
            "chart": chart_spec,
            "position": {
                "overlayPosition": {
                    "anchorCell": anchor_cell,
                    "widthPixels": 600,
                    "heightPixels": 400
                }
            }
        }
    }

    return chart_request

def add_chart_to_worksheet(gc, spreadsheet_id, sheet_id, chart_title, chart_type, data_range, anchor_cell):
    request = create_chart_request(sheet_id, chart_title, chart_type, data_range, anchor_cell)

    batch_update_request = {
        'requests': [request]
    }

    gc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_request).execute()

# Use your service account credentials
creds = service_account.Credentials.from_service_account_file('')
gc = build('sheets', 'v4', credentials=creds)

# Replace with your actual spreadsheet ID and sheet title
spreadsheet_id = "your_spreadsheet_id"
sheet_title = "Sheet1"

# Specify the data range for the chart (e.g., 'Sheet1!A1:B10')
data_range = f"{sheet_title}!A1:B{len(CategTot) + 1}"

# Specify the anchor cell for the chart
anchor_cell = 'A1'

# Specify the sheet title or adjust as needed
sheet_title = "Sheet1"

# Execute the financial management function
rows, max_category, start_row, worksheet = finmgmt(file, sheet_title)


# Print total money spent per month
# print_monTot()
#
# # Insert sum of expenses month-wise to the worksheet
# insert_monTot_in(start_row, worksheet)
#
# # Insert the category with the maximum sum to the worksheet
# insert_max_category_in(max_category, start_row, worksheet)
#
# # Insert the biggest transactions in each month to the worksheet
# insert_biggestTs_in(start_row, worksheet)
#
# # Insert net amounts to the worksheet
# insert_net_amounts_in(start_row, worksheet)

# Add a bar chart to the worksheet
add_chart_to_worksheet(gc, spreadsheet_id, sheet_title, 'Category vs Amount', 'BAR', data_range, anchor_cell)
