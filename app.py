from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    page = ''

    while (True):
        # Call the Drive v3 API
        results = service.files().list(
            pageSize=100, pageToken=page, fields="nextPageToken, files(name, permissions)").execute()
        items = results.get('files', [])
        #Breaking the loop if all files are aquired
        if len(items) < 100:
            break
        page = results.get('nextPageToken')

        details = []           # For permission details
        noInfo = []            # To include the name of the file if permission details are not available

        if not items:
            print('No files found.')
        else:
            print('Files:')

            for item in items:
                details.clear()
                details.append(item['name'])
                try:
                    for element in item['permissions']:
                        for key, value in element.items():
                            if key == 'displayName':
                                details.append(value)
                            if key == 'emailAddress':
                                details.append(value)
                            if key == 'role':
                                details.append(value)
                    print(item['name'], details)
                    with open('permission.csv', 'a', encoding='utf-8') as writeFile:
                        writer = csv.writer(writeFile)
                        writer.writerow(details)

                except:
                    noInfo.clear()
                    noInfo.append(item['name'])
                    noInfo.append("No permission information")
                    with open('permission.csv', 'a', encoding='utf-8') as writeFile:
                        writer = csv.writer(writeFile)
                        writer.writerow(noInfo)
                    print(item['name'], "No permission information")

if __name__ == '__main__':
    main()