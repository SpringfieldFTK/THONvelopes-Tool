import sys
from googleapiclient.errors import HttpError

import Log
from . import SHEET_SERVICE, DRIVE_SERVICE, FILES, PROPERTIES, TEMPLATE
import Info


def createFile(title):
    data = {
        'name': title,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [PROPERTIES['files']]
    }
    request = DRIVE_SERVICE.files().create(body=data)

    try:
        file = request.execute()
        FILES.append(file)

        update = {
            "requests": []
        }

        for sheet in TEMPLATE['sheets']:
            update['requests'].append({
                "addSheet": {
                    "properties": {
                        "title": sheet['title']
                    }
                }
            })
        update['requests'].append({
            "deleteSheet": {
                "sheetId": 0
            }
        })
        sheet_data = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=file['id'], body=update).execute()

        values_body = {
            "valueInputOption": "USER_ENTERED",
            "data": []
        }

        for sheet in TEMPLATE['sheets']:
            values_body['data'].append({
                "range": "'{0}'!A{1}:{2}{1}".format(sheet['title'], sheet['header_row'],
                                                    Info.index_to_col(len(sheet['header_columns']) - 1)),
                "values": [sheet['header_columns']]
            })
        SHEET_SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=file['id'], body=values_body).execute()

        update = {
            "requests": []
        }
        for index, sheet in enumerate(TEMPLATE['sheets']):
            update['requests'].append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_data['replies'][index]['addSheet']['properties']['sheetId'],
                        "startRowIndex": sheet['header_row'] - 1,
                        "endRowIndex": sheet['header_row']
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment)"
                }
            })
            update['requests'].append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_data['replies'][index]['addSheet']['properties']['sheetId'],
                        "gridProperties": {
                            "frozenRowCount": sheet['header_row']
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            })
            # Auto-Size Files Pros- Some columns fit better, Cons - Some columns are really, really small
            # update['requests'].append({
            #   "autoResizeDimensions": {
            #     "dimensions": {
            #       "sheetId": sheet_data['replies'][index]['addSheet']['properties']['sheetId'],
            #       "dimension": "COLUMNS",
            #       "startIndex": 0,
            #       "endIndex": len(sheet['header_columns'])
            #     }
            #   }
            # })
        SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=file['id'], body=update).execute()

    except HttpError as err:
        Log.err(error=err, raw=True)


def batchCreateFile(file, suffix=""):
    try:
        with open(file) as titles:
            for index, title in enumerate(titles):
                title = title.strip()
                if len(suffix) > 0:
                    title = title + " - " + suffix
                createFile(title)
                sys.stdout.write("\rCreated file {0}".format(index + 1))
                sys.stdout.flush()
            sys.stdout.write("\rDone\n")
            sys.stdout.flush()
    except FileNotFoundError:
        Log.err(error="File '{0}' not found".format(file))


def deleteFile(title):
    for file in FILES:
        if file['name'].lower() == title.lower():
            request = DRIVE_SERVICE.files().delete(fileId=file['id'])
            try:
                request.execute()
                FILES.remove(file)
            except HttpError as err:
                Log.err(error=err,raw=True)
            return
    Log.err(error="Spreadsheet '{0}' not found".format(title))