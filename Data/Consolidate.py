from Requests import SHEET_SERVICE, FILES, DRIVE_SERVICE

import Info


def consolidateData(origin_sheet, destination_sheet, file_id=None, parent_id=None, file_title=None, append=False):
    data = []
    first_names = ['Member First Name']
    last_names = ['Member Last Name']
    full_names = ['Member Full Name']

    if append:
        response = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=file_id, majorDimension="ROWS",
                                                             range="'{0}'!A1:Q1".format(destination_sheet)).execute()
        for cell in response['values'][0]:
            data.append([cell])

    for file in FILES:
        full_name = file['name'].split('-')[0].strip()
        last_name = full_name.split(' ')[-1]
        first_name = ' '.join(full_name.split(' ')[:-1])

        response = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=file['id'], majorDimension="ROWS", range="'{0}'!A1:H".format(origin_sheet)).execute()
        if len(data) == 0:
            for cell in response['values'][0]:
                data.append([cell])

        headings = {}
        rows = 0
        for column in data:
            for index, cell in enumerate(response['values'][0]):
                if column[0] == cell:
                    headings[cell] = index
                    break
        for row in response['values'][1:]:
            if len(row) >= (len(data)) and row[0] != "Add new addresses below this line so we can keep track of incentives!":
                for column in data:
                    a = column[0]
                    b = headings[a]
                    c = row[b]
                    column.append(c)
                rows += 1
        print("File")
        first_names.extend([first_name for x in range(rows)])
        last_names.extend([last_name for x in range(rows)])
        full_names.extend([full_name for x in range(rows)])

    data.append(first_names)
    data.append(last_names)
    data.append(full_names)
    body = {
        "values": data,
        "majorDimension": "COLUMNS"
    }

    if not file_id:
        data = {
            'name': file_title,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [parent_id]
        }
        file = DRIVE_SERVICE.files().create(body=data).execute()
        file_id = file['id']

    if not append:
        add_sheet = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": destination_sheet
                        }
                    }
                }
            ]
        }
        SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=file_id, body=add_sheet).execute()

    SHEET_SERVICE.spreadsheets().values().append(spreadsheetId=file_id, range="'{0}'!A1".format(destination_sheet), body=body, valueInputOption="USER_ENTERED").execute()
    sheet_id = Info.getSheetId({'id': file_id}, destination_sheet)
    requests = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1
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
            },
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {
                            "frozenRowCount": 1
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 11
                    }
                }
            }
        ]
    }
    if not append:
        # noinspection PyTypeChecker
        requests['requests'].append({
                "deleteSheet": {
                    "sheetId": 0
                }
            })
    SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=file_id, body=requests).execute()