import requests
import webbrowser
import openpyxl
import os
from datetime import datetime

# OAuth2 credentials - Replace with your actual credentials
CLIENT_ID = '4e97513dc1aee76830dd43ba617eaccf97f87bef'
CLIENT_SECRET = 'bd5eee02684a8e2456b3e3d94ed0b5f914c8bf29'
REDIRECT_URI = 'https://localhost:5000'

# Deputy API endpoints
BASE_URL = 'https://lakeshomes.na.deputy.com/api/v1'
AUTHORIZE_URL = 'https://once.deputy.com/my/oauth/login'
TOKEN_URL = 'https://once.deputy.com/my/oauth/access_token'

def get_access_token():
    authorize_params = {
        'client_id': CLIENT_ID,
        'redirect_uri': REDIRECT_URI,
        'response_type': 'code',
        'scope': 'longlife_refresh_token',
    }


    # Automatically open the Deputy authentication URL in the default web browser
    auth_url = f"{AUTHORIZE_URL}?client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}&response_type=code&scope=longlife_refresh_token"
    webbrowser.open(auth_url)

    # Prompt the user to enter the code obtained from the redirect URL
    code = input("Please enter the code from the redirect URL: ")

    # Exchange the code for an access token
    token_data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code',
        'code': code,
        'scope': 'longlife_refresh_token',
    }

    response = requests.post(TOKEN_URL, data=token_data)

    if response.status_code == 200:
        return response.json()['access_token']
    else:
        print("Failed to get access token.")
        return None
    
def get_all_users(access_token):
    url = f"{BASE_URL}/resource/Employee/QUERY"
    headers = {"Authorization": f"Bearer {access_token}"}

    payload = {
        "fields": "Id,DisplayName,FirstName,LastName",
        "include_inactive": True  # Set this to True to include inactive employees
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get users. Error: {response.status_code}")
        return None
    
def get_timesheets_hours(access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    url = f'{BASE_URL}/resource/Timesheet/QUERY'
    params = {
        'aggr': {"TotalTime": "sum"},
        'group': ["Employee"],
    }

    response = requests.post(url, headers=headers, json=params)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get timesheets. Error: {response.status_code}")
        return None

def get_employee_by_id(access_token, user_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
    }
    url = f'{BASE_URL}/resource/Employee/QUERY'
    params = {
        'search': {
            's1': {
                'field': 'Id',
                'data': user_id,
                'type': 'eq',
            }
        }
    }

    response = requests.post(url, headers=headers, json=params)

    if response.status_code == 200:
        employees_data = response.json()
        for employee in employees_data:
            if employee['Id'] == user_id:
                return employee
        print(f"Employee {user_id} not found.")
        return None
    else:
        print(f"Failed to get employee {user_id}. Error: {response.status_code}")
        return None

def get_timesheets_by_operational_unit(access_token, operational_unit_name, start_date, end_date):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    url = f'{BASE_URL}/resource/Timesheet/QUERY'
    params = {
        'search': {
            's1': {
                'field': 'OperationalUnitName',
                'data': operational_unit_name,
                'type': 'eq',
            },
            's2': {
                'field': 'Date',
                'data': start_date,
                'type': 'gt',
            },
            's3': {
                'field': 'Date',
                'data': end_date,
                'type': 'lt',
            },
        },
        'fields': 'Employee,TotalTime',
    }

    response = requests.post(url, headers=headers, json=params)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get timesheets for operational unit {operational_unit_name}. Error: {response.status_code}")
        return None
    
# Additional function to get timesheets for a specific date range
def get_timesheets_by_date_range(access_token, user_id, start_date, end_date):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    url = f'{BASE_URL}/resource/Timesheet/QUERY'
    params = {
        'search': {
            's1': {
                'field': 'Employee',
                'data': user_id,
                'type': 'eq',
            },
            's2': {
                'field': 'Date',
                'data': start_date,
                'type': 'gt',
            },
            's3': {
                'field': 'Date',
                'data': end_date,
                'type': 'lt',
            },
        },
        'fields': 'Id,Employee,StartTime,EndTime,TotalTime,DisplayName',  # Include the DisplayName field
    }

    response = requests.post(url, headers=headers, json=params)

   # print("Response Content:", response.text)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get timesheets for {user_id}. Error: {response.status_code}")
        return None
    
def get_employee_name_and_area(access_token, user_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    url = f'{BASE_URL}/resource/OperationalUnit/INFO'

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        fields = data.get('fields', {})

        if 'OperationalUnitName' in fields:
            employee_url = f'{BASE_URL}/resource/Employee/{user_id}'
            employee_response = requests.get(employee_url, headers=headers)

            if employee_response.status_code == 200:
                employee_data = employee_response.json()
                display_name = employee_data.get('DisplayName', '')
                area_data = employee_data.get('Areas', [])

                area_name = ''

                if isinstance(area_data, list) and area_data:
                    for area_item in area_data:
                        area_name += area_item.get('OperationalUnitName', '') + ', '

                    # Remove trailing comma and space
                    area_name = area_name.rstrip(', ')

                return display_name, area_name
            else:
                print(f"Failed to get employee {user_id}. Error: {employee_response.status_code}")
        else:
            print("OperationalUnitName is not available in the API response.")
    else:
        print(f"Failed to get OperationalUnit INFO. Error: {response.status_code}")

    return '', ''


def write_to_excel(sleep_timesheets_data, other_timesheets_data, access_token, start_date, end_date):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Timesheets Hours'

    # Write the header row
    header = ['Index', 'User ID', 'Display Name', 'Total Sleep Time', 'Total Other Time']
    sheet.append(header)

    # Create dictionaries to store sleep and other time totals per employee
    sleep_totals = {}
    other_totals = {}

    # Process sleep timesheets data
    for idx, timesheet in enumerate(sleep_timesheets_data):
        user_id = timesheet.get('Employee', '')
        display_name, area = get_employee_name_and_area(access_token, user_id)
        total_time = timesheet['TotalTime']
        
 # Categorize timesheets as sleep or other based on area
        if area == 'Sleep':
            sleep_totals[user_id] = sleep_totals.get(user_id, 0) + total_time
        else:
            other_totals[user_id] = other_totals.get(user_id, 0) + total_time

    # Process other timesheets data
    for idx, timesheet in enumerate(other_timesheets_data):
        user_id = timesheet.get('Employee', '')
        display_name, area = get_employee_name_and_area(access_token, user_id)
        total_time = timesheet['TotalTime']

        # Categorize timesheets as sleep or other based on area
        if area != 'Sleep':
            other_totals[user_id] = other_totals.get(user_id, 0) + total_time

    # Write sleep time totals for each employee
    for idx, user_id in enumerate(sleep_totals):
        display_name, _ = get_employee_name_and_area(access_token, user_id)
        sleep_time = sleep_totals.get(user_id, 0)

        # Write the data to the Excel worksheet
        row = [idx + 1, user_id, display_name, sleep_time, 0]  # Other time is 0 for sleep timesheets
        sheet.append(row)

    # Write other time totals for each employee
    for idx, user_id in enumerate(other_totals):
        display_name, _ = get_employee_name_and_area(access_token, user_id)
        other_time = other_totals.get(user_id, 0)

        # Write the data to the Excel worksheet
        row = [idx + 1, user_id, display_name, 0, other_time]  # Sleep time is 0 for other timesheets
        sheet.append(row)

    workbook.save('timesheets_totals.xlsx')
    print('Data written to Excel successfully.')

def main():
    # Step 1: Get the access token using OAuth2
    access_token = get_access_token()

    if access_token:
        # Step 2: Get all users
        users_data = get_all_users(access_token)

        if users_data:
            sleep_timesheets_data = []
            other_timesheets_data = []

            # Get user input for date range
            start_date = input("Enter the start date (YYYY-MM-DD): ")
            end_date = input("Enter the end date (YYYY-MM-DD): ")

            for user in users_data:
                user_id = user['Id']
                timesheets_data = get_timesheets_by_date_range(access_token, user_id, start_date, end_date)
                if timesheets_data:
                    user_operational_unit = user.get('OperationalUnitName', '')  # Replace with the actual field name
                    total_time = sum(float(ts['TotalTime']) for ts in timesheets_data)

                    if user_operational_unit == 'Sleep':
                        sleep_timesheets_data.append({'Employee': user_id, 'TotalTime': total_time})
                    else:
                        other_timesheets_data.append({'Employee': user_id, 'TotalTime': total_time})
            print("Sleep Timesheets Data:", sleep_timesheets_data)
            print("Other Timesheets Data:", other_timesheets_data)
            # Step 3: Write sleep and other timesheets data to Excel
            write_to_excel(sleep_timesheets_data, other_timesheets_data, access_token, start_date, end_date)

if __name__ == '__main__':
    main()

