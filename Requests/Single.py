import sys
from googleapiclient.errors import HttpError

import Log
import visuals
from . import SHEET_SERVICE, DRIVE_SERVICE, FILES
import Template
import Info
import Options


def createFile(title):
    data = {
        'name': title,
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [Options.get_indiv_folder()]
    }
    request = DRIVE_SERVICE.files().create(body=data)

    try:
        file = request.execute()
        FILES.append(file)

        update = {
            "requests": []
        }

        for title in Template.get_sheets():
            update['requests'].append({
                "addSheet": {
                    "properties": {
                        "title": title
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

        for title in Template.get_sheets():
            values_body['data'].append({
                "range": "'{0}'!A{1}:{2}{1}".format(title, Template.get_header_row(title),
                                                    Info.index_to_col(len(Template.get_columns(title)) - 1)),
                "values": [Template.get_columns(title)]
            })
        SHEET_SERVICE.spreadsheets().values().batchUpdate(spreadsheetId=file['id'], body=values_body).execute()

        update = {
            "requests": []
        }
        for index, title in enumerate(Template.get_sheets()):
            update['requests'].append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_data['replies'][index]['addSheet']['properties']['sheetId'],
                        "startRowIndex": Template.get_header_row(title) - 1,
                        "endRowIndex": Template.get_header_row(title)
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
                            "frozenRowCount": Template.get_header_row(title)
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


def batchCreateFile(files, suffix=""):
    for index, title in enumerate(files):
        title = title.strip()
        if suffix:
            title = title + " - " + suffix
        createFile(title)
        sys.stdout.write("\rCreated file {0}".format(index + 1))
        sys.stdout.flush()
        sys.stdout.write("\rDone\n")
        sys.stdout.flush()
        visuals.set_progress(index+1, len(files))


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