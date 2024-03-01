from datetime import datetime
import win32com.client
import os
import time


def refresh_excel_connections(file_name):
    # Get the current working directory
    current_directory = os.getcwd()
    # Construct the full path to the Excel file
    excel_file_path = os.path.join(current_directory, file_name)
    # Create Excel Application object
    excel_app = win32com.client.Dispatch("Excel.Application")
    # Open the Excel workbook
    workbook = excel_app.Workbooks.Open(excel_file_path)
    # Refresh all connections in the workbook
    print("Refreshing Excel Sheet...\n")
    workbook.RefreshAll()
    time.sleep(10)
    # Save the changes
    workbook.Save()
    # Close the workbook
    workbook.Close()

    # Quit Excel application
    excel_app.Quit()

def process_datetime(global_variables):
    # Format the DateTime value as required
    
    #format_date_datetime = datetime.strptime(global_variables.FormatDate, '%Y-%m-%d %H:%M:%S')
    if 0 <= global_variables.FormatDate.weekday() <= 4 and 8 <= global_variables.FormatDate.hour < 17:
        global_variables.Billing = 'standard'
    else:
        global_variables.Billing = 'outside hours'

    # Format the DateTime value as required
    global_variables.FormatDate = global_variables.FormatDate.strftime('%Y/%m/%dT%H:%M:%S')

