import json
import requests
import openpyxl

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


def send_post_request(request_data, variable_value=None):
    try:
        # Parse the JSON request data
        request_json = json.loads(request_data)

        # Extract parameters from the JSON data
        uri = request_json.get("uri")
        method = request_json.get("method", "POST")
        headers = request_json.get("headers", {})
        authentication = request_json.get("authentication", {})
        body = request_json.get("body", "")

        # Replace the variable in the body with the most recent entry from the Excel file
        if variable_value is None:
            variable_value = get_most_recent_entry("../../../Dropbox (CSU Fullerton)/Admin and Conference Services/Technology/Lockouts/Lockouts.xlsx", "Lockouts List", "A")

        # Prepare the authentication details
        auth_type = authentication.get("type", "")
        username = authentication.get("username", "")
        password = authentication.get("password", "")

        # Prepare the request headers
        headers["Content-Type"] = "application/json"

        # Prepare the request
        response = requests.request(
            method,
            uri,
            headers=headers,
            auth=(username, password) if auth_type.lower() == "basic" else None,
            data=body.replace('{variable}', str(variable_value)) if body else None

        )

        # Print the response
        print("Response Code:", response.status_code)
        print("Response Content:", response.text)

    except json.JSONDecodeError as e:
        print("Error decoding JSON:", str(e))
    except Exception as e:
        print("An error occurred:", str(e))

if __name__ == "__main__":
    # Example JSON data with a variable placeholder '{cwid}'
    example_json = '''
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
    #variable_value = get_most_recent_entry("Lockouts.xlsx", "Lockouts List", "A")

    # Send the POST request using the example JSON data and variable value
    send_post_request(example_json)
