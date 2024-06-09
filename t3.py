import gspread
from oauth2client.service_account import ServiceAccountCredentials
import matplotlib.pyplot as plt
import pygsheets

# Set the scope and credentials for Google Sheets API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

# Open the Google Sheet by title or URL
# Replace 'Your Google Sheet Name' with your actual sheet name
worksheet = gc.open('pyPt1').sheet1

# Get data from the sheet
data = worksheet.get_all_values()

# Extract columns for x and y values
x_values = []
y_values = []

# Extract columns for x and y values
x_labels = [row[0] for row in data[1:6]]  # Assuming string data is in the first column (A2 to A6)
y_values = [float(row[2]) for row in data[1:6]]  # Assuming float data is in the third column (C2 to C6)

# Plot the data
plt.plot(x_labels, y_values, marker='o')  # Using markers for each data point
plt.xlabel('Dates')
plt.ylabel('Month Total')
plt.title('Monthwise Expense')


# Save the plot as an image (optional)
plt.savefig('graph.png')

# Show the plot in the console (optional)
plt.show()
