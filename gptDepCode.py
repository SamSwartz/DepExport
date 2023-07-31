import requests
import webbrowser
import openpyxl

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
    
def get_all_users():
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
    
#def get_timesheets_hours(access_token, user_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }
    url = f'{BASE_URL}/resource/TimesheetPayReturn/QUERY'
    params = {
        'user': user_id,
    }

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to get timesheets for user {user_id}. Error: {response.status_code}")
        return None

def write_to_excel(data):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Timesheets Hours'

    # Assuming the response contains a list of timesheets with hours data
    # Replace this with the actual structure of the API response
    for idx, timesheet in enumerate(data):
        # Assuming the structure of the timesheet data
        # Replace these keys with actual keys from the API response
        user_id = timesheet['user_id']
        start_time = timesheet['start_time']
        end_time = timesheet['end_time']
        total_hours = timesheet['total_hours']

        # Write the data to the Excel worksheet
        row = [idx + 1, user_id, start_time, end_time, total_hours]
        sheet.append(row)

    workbook.save('timesheets_hours.xlsx')
    print('Data written to Excel successfully.')

def main():
    # Replace 'user_id' with the actual user ID for whom you want to retrieve timesheets
    user_id = 'user_id'

    # Step 1: Get the access token using OAuth2
    access_token = get_access_token()

    if access_token:
        # Step 2: Get all users
        users_data = get_all_users(access_token)

    if users_data:
        all_timesheets_data = []
        for user in users_data:
            user_id = user['Id']
            timesheets_data = get_timesheets_hours(access_token, user_id)
            if timesheets_data:
                all_timesheets_data.extend(timesheets_data)

        # Step 3: Write all timesheets data to Excel
        write_to_excel(all_timesheets_data)

if __name__ == '__main__':
    main()