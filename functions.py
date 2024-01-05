import json
import  jsonschema
import requests
import openpyxl
from datetime import datetime

class GlobalVariables:
    CWID = ''
    EntryID=''
    NamePreferred = ''
    NameFirst = 'test'
    NameLast = ''
    EntryID = ''
    RoomSpaceID = ''
    Description = ''
    Amount = 0
    LockoutCount = 100
    cwid_list = []
    time_list = []
    index = 0
    TermID = 103
    bookingDataBody = "select roomspaceid, [roomspace].description from booking where entryid='{variable}' and entrystatusenum IN (2, 5)"
    cwidLookupBody = "select namepreferred, namefirst, namelast, entryid from entry where id1='{variable}'"
    CwidcountURI = "https://fullerton.starrezhousing.com/StarRezREST/services/getreport/7232.json?entryID={variable}&termID=103"
    AddGenericDataURI = "https://fullerton.starrezhousing.com/StarRezREST/services/create/lockouts/"
    AddGenericDataBody = '''
    {
        "ItemID": "{variable1}",
        "TimeStamp": "{variable2}",
        "RoomSpaceID": "{variable3}",
        "TermID": 103,
        "Username": "Housing@fullerton.edu",
        "Processed":"False"
    }
    '''
    data_json = '''
    {
        "uri": "https://fullerton.starrezhousing.com/StarRezREST/services/query",
        "method": "POST",
        "headers": {
            "Accept": "application/json"
        },
        "authentication": {
            "username": "housingrest@fullerton.edu",
            "password": "1bfb9011-6989-4d34-9eec-238e224ab253",
            "type": "Basic"
        },
        "body": "select namepreferred, namefirst, namelast, entryid from entry where id1='{variable}'"
    }
    '''


#***HELPER FUNCTIONS***
#----------------------------------------------------------------------------------------------------------------
def process_datetime(input_datetime_str):
    # Convert input string to datetime object
    input_datetime = datetime.strptime(input_datetime_str, '%d/%m/%YT%H:%M:%S')

    # Check if it's a weekday between 8am and 5pm
    if 0 <= input_datetime.weekday() <= 4 and 8 <= input_datetime.hour < 17:
        result_variable = 10
    else:
        result_variable = 20

    return result_variable

# Example usage:
    input_datetime_str = "12/20/2023 15:11"
    result_value = process_datetime(input_datetime_str)
    print(f"The result variable is set to: {result_value}")



#----------------------------------------------------------------------------------------------------------------

def get_unprocessed():
    # Load the Excel workbook
    workbook = openpyxl.load_workbook('Dropbox (CSU Fullerton)/Admin and Conference Services/Technology/Lockouts/Lockouts.xlsx')

    # Assuming you are working with the first sheet (you can change it as needed)
    sheet = workbook['Lockouts List']

    # Specify the columns you are interested in
    column_with_value = 'A'  # Change this to the column letter you are interested in
    column_with_value2 = 'B'
    column_to_check = 'F'  # Change this to the column letter where you want to check for the value

    # Iterate through the rows
    # Iterate through the rows starting from the third row
    for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        # Assuming headers are in the first row, so starting from the third row (min_row=3)
        cell_value = row[sheet[column_to_check + '1'].column - 1]  # Adjust index for 0-based Python index
        if cell_value != "yes":
            GlobalVariables.cwid_list.append(row[sheet[column_with_value + '1'].column - 1])
            date_time_value = row[sheet[column_with_value2 + '1'].column - 1]

            # Format the DateTime value as required
            formatted_datetime = datetime.strftime(date_time_value, '%Y/%m/%dT%H:%M:%S')
            # Append the formatted datetime to the list
            GlobalVariables.time_list.append(formatted_datetime)


            # Update the corresponding cell in the 'processed' column for the current row
            #sheet[column_to_check + str(row_index)].value = "yes"
            # Save the changes to the workbook
    workbook.save('Dropbox (CSU Fullerton)/Admin and Conference Services/Technology/Lockouts/Lockouts.xlsx')

    # Close the workbook when done
    workbook.close()


#----------------------------------------------------------------------------------------------------------------


# not used but possibly helpful function for testing
def get_most_recent_entry(file_path, sheet_name, column_name):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]

        # Assuming the entries are in column 'A' and have headers in the first row
        column_values = [cell.value for cell in sheet[column_name] if cell.value is not None]

        # Get the most recent entry
        most_recent_entry = column_values[-1] if column_values else None

        return most_recent_entry

    except Exception as e:
        print("Error reading Excel file:", str(e))
        return None

#----------------------------------------------------------------------------------------------------------------
def process_entry(entry,state):

    # Extract specific fields into separate variables
    if state == 0:
        GlobalVariables.NamePreferred = entry.get("NamePreferred")
        GlobalVariables.NameFirst = entry.get("NameFirst")
        GlobalVariables.NameLast = entry.get("NameLast")
        GlobalVariables.EntryID = entry.get("EntryID")
    elif state == 1:
        GlobalVariables.RoomSpaceID = entry.get("RoomSpaceID")
        GlobalVariables.Description = entry.get("Description")
    elif state == 2:
        GlobalVariables.LockoutCount = entry.get("Count")



#----------------------------------------------------------------------------------------------------------------

def send_post_request(request_data, state, index):
    variable_value = ''

    # Replace the variable in the body with the most recent entry from the Excel file
    if state==0:
        variable_value = GlobalVariables.CWID
    elif state==1:
        variable_value = GlobalVariables.EntryID
    elif state == 2:
        variable_value = GlobalVariables.EntryID
        json_dict=json.loads(request_data)
        json_dict["uri"] = GlobalVariables.CwidcountURI
        json_dict["body"] = ''
        json_dict["method"] = 'GET'
        request_data=json.dumps(json_dict)
    elif state == 3:
        variable1 = GlobalVariables.EntryID
        json_dict=json.loads(request_data)
        json_dict["uri"] = GlobalVariables.AddGenericDataURI
        json_dict["body"] = GlobalVariables.AddGenericDataBody
        request_data = json.dumps(json_dict)

    try:
        # Parse the JSON request data
        request_json = json.loads(request_data)

        # Extract parameters from the JSON data
        uri = request_json.get("uri")
        method = request_json.get("method")
        headers = request_json.get("headers", {})
        authentication = request_json.get("authentication", {})
        body = request_json.get("body", "")

        # Prepare the authentication details
        auth_type = authentication.get("type", "")
        username = authentication.get("username", "")
        password = authentication.get("password", "")

        # Prepare the request headers
        headers["Content-Type"] = "application/json"

        if state == 2:
            uri = uri.replace('{variable}', str(variable_value))
        if state == 3:
            body = body.replace('{variable1}', str(GlobalVariables.EntryID))
            body = body.replace('{variable2}', str(GlobalVariables.time_list[index]))
            body = body.replace('{variable3}', str(GlobalVariables.RoomSpaceID))
            #body = body.replace('{variable4}', GlobalVariables.TermID)

            print(body)
        # Prepare the request
        response = requests.request(
            method,
            uri,
            headers=headers,
            auth=(username, password) if auth_type.lower() == "basic" else None,
            data=body if state == 2 else body.replace('{variable}', str(variable_value)) if body else None
            )


        # Store the received data as a JSON object
        received_data = response.json()
        if state==3:
            print (response.text)
        #parse json
        for entry in response.json():
            process_entry(entry,state)


    except json.JSONDecodeError as e:
        print("Error decoding JSON:", str(e))
    except Exception as e:
        print("An error occurred:", str(e))

#----------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------


#***DRIVER FUNCTIONS***

#this function facilitates looking up students data from CWID and adding a lockout
#record to generic Data
def AddGenericData():
    GlobalVariables()
    state = 0

    #retrieve list of unprocessed cwids
    get_unprocessed()

    #loop that tracks state of the process
    for temp in GlobalVariables.cwid_list:
        GlobalVariables.CWID = temp
        if state==0:
            json_dict=json.loads(GlobalVariables.data_json)
            json_dict["body"] = GlobalVariables.cwidLookupBody
            GlobalVariables.data_json=json.dumps(json_dict)
            send_post_request(GlobalVariables.data_json, state, GlobalVariables.index)
            state+=1
        if state==1:
            # Update the existing JSON object with the new body value
            json_dict=json.loads(GlobalVariables.data_json)
            json_dict["body"] = GlobalVariables.bookingDataBody
            GlobalVariables.data_json=json.dumps(json_dict)
            send_post_request(GlobalVariables.data_json, state, GlobalVariables.index)
            state+=1
        if state == 2:
            send_post_request(GlobalVariables.data_json, state, GlobalVariables.index)
            state += 1
        if state==3:
            send_post_request(GlobalVariables.data_json, state, GlobalVariables.index)
            state=0

        GlobalVariables.index+=1
        print(GlobalVariables.NameFirst + '\n' + GlobalVariables.NameLast + '\n' + str(GlobalVariables.CWID) + '\n' + str(GlobalVariables.EntryID) + '\n' + str(GlobalVariables.RoomSpaceID) + '\n' + GlobalVariables.Description + '\n' + str(GlobalVariables.LockoutCount) + '\n')
