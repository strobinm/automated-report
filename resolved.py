import datetime as dt
import xlwings as xw
import os


# Connect to the active Excel application
app = xw.apps.active
# Turn off screen updating
app.screen_updating = False
PSHandover = xw.Book('ICQA PS HANDOVER testy.xlsm')
handoverSheet = PSHandover.sheets['PS Handover']
# get Shift_Start_Date value
date = str(handoverSheet['B1'].options(dates=dt.date).value)
temp_Excel_File_Name = 'Bin Item Defects All types ' + date + ".csv"
user = handoverSheet['I4'].value
# Open the workbook in the Downloads directory
wb = xw.Book(r'C:\Users\{}\Downloads\{}'.format(user, temp_Excel_File_Name))
# Select the active sheet
sht = wb.sheets.active
# Get the last row in the sheet
last_row = sht.cells.last_cell.row
# Create Shift_Start_datetime filter
shift = handoverSheet['F1'].value
if (shift == 'DS'):
    date_time_filter = dt.datetime.strptime(
        date, "%Y-%m-%d").strftime('%d/%m/%Y') + ', 06:20:00'
elif (shift == 'NS'):
    date_time_filter = dt.datetime.strptime(
        date, "%Y-%m-%d").strftime('%d/%m/%Y') + ', 17:50:00'
# Create a list to store the selected rows
selected_rows = list(filter(lambda c: c[7] == user and c[10] > date_time_filter, sht.range(
    'A2:K' + str(last_row)).value))
# Create a dictionary to count the frequency of each type
frequency = {}
for col in selected_rows:
    type = col[2]
    if type not in frequency:
        frequency[type] = 0
    frequency[type] += 1
# Function to set the frequency values to a specified column in the matching row
for i in range(3, 17):
    type = handoverSheet.range("B" + str(i)).value
    if type in frequency:
        handoverSheet.range("D" + str(i)).value = frequency[type]
# Close the workbook
wb.close()
if os.path.exists(r'C:\Users\{}\Downloads\{}'.format(user, temp_Excel_File_Name)):
    os.remove(r'C:\Users\{}\Downloads\{}'.format(user, temp_Excel_File_Name))
else:
    print("The file does not exist.")
app.screen_updating = True
print("done")
