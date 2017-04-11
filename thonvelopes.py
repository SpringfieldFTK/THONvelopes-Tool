import httplib2
import os
import string
import re
import json

from interface import Interface
from interface import Command
from interface import Parameter

# noinspection PyUnresolvedReferences
from apiclient import discovery
# noinspection PyUnresolvedReferences
from apiclient.errors import HttpError
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from pprint import pprint

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    print("No Flags")
    flags = None

SCOPES = 'https://www.googleapis.com/auth/drive ' + 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'THONvelope Helper API'

DRIVE_SERVICE = None
SHEET_SERVICE = None
FILES = []

PROPERTIES = {
    'files': "0BzzhJ1bdYANAczU1b0tQby1iOUk",
    'header_row': 1
}


def print_error(spreadsheet, error, raw=False):
    if raw:
        m = re.search("\".+: (.+)\"", str(error))
        error = m.group(1)
    print("ERROR in {0}({1}): {2}".format(spreadsheet['name'], spreadsheet['id'], error))


def col_to_index(col):
    if type(col) is str:
        col = list(reversed(col.upper()))
    return string.ascii_uppercase.index(col.pop()) if len(col) == 1 else 26 * (
    string.ascii_uppercase.index(col.pop()) + 1) + col_to_index(col)


def index_to_col(n):
    div = n + 1
    string = ""
    while div > 0:
        module = (div - 1) % 26
        string = chr(65 + module) + string
        div = int((div - module) / 26)
    return string


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'thonvelope-helper.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
    return credentials


def getFileList():
    global FILES
    query = "'{0}' in parents and mimeType='application/vnd.google-apps.spreadsheet'".format(PROPERTIES['files'])
    results = DRIVE_SERVICE.files().list(q=query, fields='files(id, name)', orderBy='name', pageSize=20).execute()
    FILES = results.get('files', [])
    if not FILES:
        print('No individual files found')
    else:
        print("{0} individual files found".format((len(FILES))))


def addSheet(spreadsheet, title, index=-1):
    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": title
                    }
                }
            }
        ]
    }
    if int(index) >= 0:
        body['requests'][0]['addSheet']['properties']['index'] = int(index) - 1
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body=body)

    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def getSheetId(spreadsheet, title):
    request = SHEET_SERVICE.spreadsheets().get(spreadsheetId=spreadsheet['id'])

    try:
        response = request.execute()
        for sheet in response['sheets']:
            if sheet['properties']['title'].lower() == title.lower():
                return sheet['properties']['sheetId']
    except HttpError as err:
        print_error(spreadsheet, err, True)
    print_error(spreadsheet, "Sheet '{0}' not found".format(title))
    return -1


def getSheetIndex(spreadsheet, title):
    request = SHEET_SERVICE.spreadsheets().get(spreadsheetId=spreadsheet['id']).execute()
    try:
        response = request.execute()
        for sheet in response['sheets']:
            if sheet['properties']['title'].lower() == title.lower():
                return sheet['properties']['index']
    except HttpError as err:
        print_error(spreadsheet, err, True)
    print_error(spreadsheet, "Sheet '{0}' not found".format(title))
    return -1


def deleteSheet(spreadsheet, title):
    id = getSheetId(spreadsheet, title=title)

    if id < 0:
        return

    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "deleteSheet": {
                    "sheetId": id
                }
            }
        ]
    })

    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def renameSheet(spreadsheet, oldTitle, newTitle):
    id = getSheetId(spreadsheet, title=oldTitle)
    if id < 0:
        return
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": id,
                        "title": newTitle
                    },
                    "fields": "title"
                }
            }
        ]
    })

    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def moveSheet(spreadsheet, title, index):
    id = getSheetId(spreadsheet, title=title)
    if id < 0:
        return

    if getSheetIndex(spreadsheet, title) <= (int(index) - 1):
        index = int(index)
    else:
        index = int(index) - 1
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": id,
                        "index": index
                    },
                    "fields": "index"
                }
            }
        ]
    })

    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def getColumnIndex(spreadsheet, sheet_title, contents):
    request = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=spreadsheet['id'],
                                                        range="'{0}'!A1:ZZZ1".format(sheet_title),
                                                        majorDimension="ROWS")
    try:
        response = request.execute()
        for i, cell in enumerate(response['values'][0]):
            if cell.lower() == contents.lower():
                return i
    except HttpError as err:
        print_error(spreadsheet, err, True)

    print_error(spreadsheet, "Column '{0}' not found in '{1}'".format(contents, sheet_title))
    return -1


def getTotalColumns(spreadsheet, sheet_title):
    request = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=spreadsheet['id'],
                                                        range="'{0}'!A1:ZZZ1".format(sheet_title),
                                                        majorDimension="ROWS")
    try:
        response = request.execute()
        end = (response['range'].split(':'))[1]
        m = re.search('[A-Z]+', end)
        return col_to_index(m.group(0)) + 1
    except HttpError as err:
        print_error(spreadsheet, err, True)


def getUtilizedColumns(spreadsheet, sheet_title, rownum=1):
    request = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=spreadsheet['id'],
                                                         range="'{0}'!A{1}:ZZZ{2}".format(sheet_title, rownum, rownum),
                                                         majorDimension="ROWS")

    try:
        response = request.execute()
        if 'values' in response:
            return len(response['values'][0])
        else:
            return 0
    except HttpError as err:
        print_error(spreadsheet, err, True)

    return -1


def insertColumn(spreadsheet, sheet_title, contents, index=0):
    index = int(index)
    if index <= 0:
        index = getUtilizedColumns(spreadsheet, sheet_title)
        if index == getTotalColumns(spreadsheet, sheet_title):
            request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
                "requests": [
                    {
                        "appendDimension": {
                            "sheetId": getSheetId(spreadsheet, sheet_title),
                            "dimension": "COLUMNS",
                            "length": 1
                        }
                    }
                ]

            })
            try:
                request.execute()
            except HttpError as err:
                print_error(spreadsheet, err, True)
                return
        index += 1
    else:
        request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": getSheetId(spreadsheet, sheet_title),
                            "dimension": "COLUMNS",
                            "startIndex": index - 1,
                            "endIndex": index
                        }
                    }
                }
            ]

        })
        try:
            request.execute()
        except HttpError as err:
            print_error(spreadsheet, err, True)
            return
    values = [
        [
            contents
        ],
    ]
    body = {
        'values': values
    }
    col = index_to_col(index - 1)
    request = SHEET_SERVICE.spreadsheets().values().update(
        spreadsheetId=spreadsheet['id'], range="'{0}'!{1}1:{2}1".format(sheet_title, col, col),
        valueInputOption="USER_ENTERED", body=body)
    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def renameColumn(spreadsheet, sheet_title, old_column, new_column):
    index = getColumnIndex(spreadsheet, sheet_title, old_column)
    if index < 0:
        return
    col = index_to_col(index)
    values = [
        [
            new_column
        ],
    ]
    body = {
        'values': values
    }
    request = SHEET_SERVICE.spreadsheets().values().update(
        spreadsheetId=spreadsheet['id'], range="'{0}'!{1}1:{2}1".format(sheet_title, col, col),
        valueInputOption="USER_ENTERED", body=body)

    try:
        request.execute()
    except HttpError as err:
        print_error(spreadsheet, err, True)


def deleteColumn(spreadsheet, sheet_title, column):
    index = getColumnIndex(spreadsheet, sheet_title, column)
    if index < 0:
        return
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": getSheetId(spreadsheet, sheet_title),
                        "dimension": "COLUMNS",
                        "startIndex": index,
                        "endIndex": index + 1
                    }
                }
            },
        ]

    })
    try:
        request.execute()
    except HttpError as err:
        pprint(err)


def moveColumn(spreadsheet, sheet_title, column, location):
    location = int(location)
    index = getColumnIndex(spreadsheet, sheet_title, column)

    if index < 0:
        return

    if location <= index:
        location -= 1
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "moveDimension": {
                    "source": {
                        "sheetId": getSheetId(spreadsheet, sheet_title),
                        "dimension": "COLUMNS",
                        "startIndex": index,
                        "endIndex": index + 1
                    },
                    "destinationIndex": location
                }
            }
        ]

    })

    try:
        request.execute()
    except HttpError as err:
        pprint(err)


def make_interface():
    interface = Interface()
    interface.set_entities(FILES)
    interface.parse_json('commands.json')

    interface.get('sheets').get('add').set_function(
        lambda sheet, args: addSheet(sheet, args[0]) if len(args) == 1 else addSheet(sheet, args[0],
                                                                                     args[1])
    )
    interface.get('sheets').get('delete').set_function(lambda sheet, args: deleteSheet(sheet, args[0]))
    interface.get('sheets').get('rename').set_function(lambda sheet, args: renameSheet(sheet, args[0], args[1]))
    interface.get('sheets').get('move').set_function(lambda sheet, args: moveSheet(sheet, args[0], args[1]))

    interface.get('columns').get('insert').set_function(
        lambda sheet, args: insertColumn(sheet, args[0], args[1]) if len(args) == 2 else insertColumn(sheet,
                                                                                                      args[0],
                                                                                                      args[1],
                                                                                                      args[2])
    )
    interface.get('columns').get('delete').set_function(lambda sheet, args: deleteColumn(sheet, args[0], args[1]))
    interface.get('columns').get('rename').set_function(
        lambda sheet, args: renameColumn(sheet, args[0], args[1], args[2]))
    interface.get('columns').get('move').set_function(lambda sheet, args: moveColumn(sheet, args[0], args[1], args[2]))

    interface.run()


def main():
    global DRIVE_SERVICE, SHEET_SERVICE

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    DRIVE_SERVICE = discovery.build('drive', 'v3', http=http)
    getFileList()

    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    SHEET_SERVICE = discovery.build('sheets', 'v4', http=http,
                                    discoveryServiceUrl=discoveryUrl)

    make_interface()


if __name__ == '__main__':
    main()
