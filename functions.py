import json
import requests
import pandas as pd
import jsonschema
from HelperFunctions import *

class GlobalVariables:
    CWID = ''
    EntryID = ''
    TermID = '103'
    RoomSpaceID =''
    Description =''
    FormatDate = ""
    Billing = 'standard'
    excel_file_name = 'LockoutsBeta.xlsx'

def SendRequest(json_file_path, state):
    # Read JSON file
    with open(json_file_path, 'r') as json_file:
        request_data = json.load(json_file)

    # Extract request details from JSON
    uri = request_data.get('uri')
    method = request_data.get('method', 'POST')
    headers = request_data.get('headers', {})
    authentication = request_data.get("authentication", {})
    body = request_data.get('body', {})

    # Prepare the request headers
    headers["Content-Type"] = "application/json"

    # Authenticate if authentication details are provided
    username = authentication.get('username')
    password = authentication.get('password')
    auth = requests.auth.HTTPBasicAuth(username, password)

    if state == 0:
        body = body.replace('{variable}', str(GlobalVariables.CWID))
    if state == 1:
        body = body.replace('{variable}', str(GlobalVariables.EntryID))
    if state == 2:
        body = json.dumps(body)
        body = body.replace('{variable1}', str(GlobalVariables.EntryID))
        body = body.replace('{variable2}', str(GlobalVariables.FormatDate))
        body = body.replace('{variable3}', str(GlobalVariables.RoomSpaceID))
        body = body.replace('{variable4}', GlobalVariables.TermID)
        body = body.replace('{variable5}', str('standard'))

    #Send POST request with JSON data
    response = requests.request(
        method,
        uri,
        headers=headers,
        auth=(username, password),
        data=body
   )
    
    for entry in response.json():
        if state == 0:
            GlobalVariables.EntryID = entry.get("EntryID")
        elif state == 1:
            GlobalVariables.RoomSpaceID = entry.get("RoomSpaceID")
            GlobalVariables.Description = entry.get("Description")

def Driver():

    refresh_excel_connections(GlobalVariables.excel_file_name)
    df = pd.read_excel(GlobalVariables.excel_file_name)
    # Loop through each row
    with open('Log.txt', 'w') as file:
        for index, row in df.iterrows():
            if pd.isnull(row['CWID']):  # Check if Column A is blank
                break  # Stop if Column A is blank
            if pd.isnull(row['Processed']):  # Check if Column F is blank
                GlobalVariables.CWID = row['CWID']  # Get value from Column A
                GlobalVariables.FormatDate = (row['Completion time'])
                process_datetime(GlobalVariables)
                SendRequest('jsons/GetEntryID.json', 0)
                
                df['Processed'] = df['Processed'].astype(str)  # Convert the column to string dtype
                df.at[index, 'Processed'] = str('yes')  # Write 'yes' to Column B
                
            SendRequest('jsons/GetBooking.json', 1)
            SendRequest('jsons/AddGenericData.json', 2)
            file.write(str(GlobalVariables.CWID) + '\n' + str(GlobalVariables.EntryID) + '\n' + str(GlobalVariables.RoomSpaceID) + '\n' + GlobalVariables.Description + '\n')
            file.write("-------------------------\n")

        df.to_excel(GlobalVariables.excel_file_name, index=False)

