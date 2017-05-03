import json
from sys import flags

import httplib2

import os
from apiclient import discovery
# noinspection PyUnresolvedReferences
from apiclient.errors import HttpError
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import Options


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, tools.argparser.parse_args([]))
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def getFileList():
    global FILES
    folder = Options.get_indiv_folder()
    print(folder)
    query = "'{0}' in parents and mimeType='application/vnd.google-apps.spreadsheet'".format(folder)
    results = DRIVE_SERVICE.files().list(q=query, fields='files(id, name)', orderBy='name', pageSize=1000).execute()
    FILES = results.get('files', [])
    if not FILES:
        print('No individual files found')
    else:
        print("{0} individual files found".format((len(FILES))))

SCOPES = 'https://www.googleapis.com/auth/drive ' + 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = './Resources/client_secret.json'
APPLICATION_NAME = 'THONvelope Helper API'

FILES = []

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
DRIVE_SERVICE = discovery.build('drive', 'v3', http=http)
getFileList()

discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                'version=v4')
SHEET_SERVICE = discovery.build('sheets', 'v4', http=http,
                                discoveryServiceUrl=discoveryUrl)

