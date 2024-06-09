# Importing required library
import pygsheets

# Create the Client
client = pygsheets.authorize(service_account_file="credentials.json")

# opens a spreadsheet by its name/title
spreadsht = client.open("pyPt1")

# opens a worksheet by its name/title
worksht = spreadsht.worksheet("title", "Sheet1")

# Creating a basic bar chart
worksht.add_chart(("A2", "A6"), [("C2", "C6")], "Monthwise Expense" )
# worksht.add_chart(("A10", "A14"), [("C10", "C14")], "Max Transaction" )
worksht.add_chart(("A17", "A18"), [("E17", "E18")], "CD/DD" )


