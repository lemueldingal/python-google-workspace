import os
import pickle
import requests
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these SCOPES, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']

def Create_Service(client_secret_file, api_name, api_version, *scopes):
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    
    cred = None
    pickle_file = rf'C:\OpenRPA\creds\token_{API_SERVICE_NAME}_{API_VERSION}.pickle'

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server(port=0)

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None

def download_google_sheet_as_excel(file_id, sheet_id, token, output_file):
    url = rf'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx&gid={sheet_id}'

    headers = {
        'Authorization': f'Bearer {token}'
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        with open(output_file, 'wb') as file:
            file.write(response.content)
        print(f"File downloaded successfully and saved as {output_file}")
    else:
        print(f"Failed to download file. Status code: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    client_secret_file = r'C:\OpenRPA\creds\fga-openrpa-oauth2.json'
    FILE_ID = 'file_id'
    SHEET_ID = 'sheet_id'
    OUTPUT_FILE = rf'C:\OpenRPA\downloads\output.xlsx' # Overwrite is default

    service = Create_Service(client_secret_file, 'drive', 'v3', SCOPES)
    
    if service:
        creds = service._http.credentials
        token = creds.token
        
        download_google_sheet_as_excel(FILE_ID, SHEET_ID, token, OUTPUT_FILE)
