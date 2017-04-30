from googleapiclient.errors import HttpError

from . import SHEET_SERVICE, FILES
import Log
import Info
import visuals


def bulk(function):
    def wrapper(*args, **kwargs):
        l = len(FILES)
        for index, file in enumerate(FILES):
            kwargs['spreadsheet'] = file
            function(*args,**kwargs)
            visuals.set_progress(index+1, l)
    return wrapper


@bulk
def addSheet(title, cols, index=-1, spreadsheet=None):
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
        Log.err(spreadsheet, err, True)

    body = {
        "requests":[
            {
                "repeatCell": {
                    "range": {
                        "sheetId": Info.getSheetId(spreadsheet, title),
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
                        "sheetId": Info.getSheetId(spreadsheet, title),
                        "gridProperties": {
                            "frozenRowCount": 1
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            }
        ]
    }

    format_request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body=body)

    values = [
        cols
    ]
    body = {
        'values': values
    }
    col = Info.index_to_col(len(cols) - 1)
    value_request = SHEET_SERVICE.spreadsheets().values().update(
        spreadsheetId=spreadsheet['id'], range="'{0}'!A1:{1}1".format(title, col),
        valueInputOption="USER_ENTERED", body=body)

    try:
        format_request.execute()
        value_request.execute()
    except HttpError as err:
        Log.err(spreadsheet, err, True)


@bulk
def deleteSheet(title, spreadsheet=None):
    id = Info.getSheetId(spreadsheet, title=title)

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
        Log.err(spreadsheet, err, True)



@bulk
def renameSheet(spreadsheet, oldTitle, newTitle):
    id = Info.getSheetId(spreadsheet, title=oldTitle)
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
        Log.err(spreadsheet, err, True)


@bulk
def moveSheet(spreadsheet, title, index):
    id = Info.getSheetId(spreadsheet, title=title)
    if id < 0:
        return

    if Info.getSheetIndex(spreadsheet, title) <= (int(index) - 1):
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
        Log.err(spreadsheet, err, True)


@bulk
def insertColumn(spreadsheet, sheet_title, contents, index=0):
    index = int(index)
    if index <= 0:
        index = Info.getUtilizedColumns(spreadsheet, sheet_title)
        if index == Info.getTotalColumns(spreadsheet, sheet_title):
            request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
                "requests": [
                    {
                        "appendDimension": {
                            "sheetId": Info.getSheetId(spreadsheet, sheet_title),
                            "dimension": "COLUMNS",
                            "length": 1
                        }
                    }
                ]

            })
            try:
                request.execute()
            except HttpError as err:
                Log.err(spreadsheet, err, True)
                return
        index += 1
    else:
        request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": Info.getSheetId(spreadsheet, sheet_title),
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
            Log.err(spreadsheet, err, True)
            return
    values = [
        [
            contents
        ],
    ]
    body = {
        'values': values
    }
    col = Info.index_to_col(index - 1)
    request = SHEET_SERVICE.spreadsheets().values().update(
        spreadsheetId=spreadsheet['id'], range="'{0}'!{1}1:{2}1".format(sheet_title, col, col),
        valueInputOption="USER_ENTERED", body=body)
    try:
        request.execute()
    except HttpError as err:
        Log.err(spreadsheet, err, True)


@bulk
def renameColumn(spreadsheet, sheet_title, old_column, new_column):
    index = Info.getColumnIndex(spreadsheet, sheet_title, old_column)
    if index < 0:
        return
    col = Info.index_to_col(index)
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
        Log.err(spreadsheet, err, True)


@bulk
def deleteColumn(spreadsheet, sheet_title, column):
    index = Log.err(spreadsheet, sheet_title, column)
    if index < 0:
        return
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": Info.getSheetId(spreadsheet, sheet_title),
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
        Log.err(spreadsheet, err, True)

@bulk
def moveColumn(spreadsheet, sheet_title, column, location):
    location = int(location)
    index = Info.getColumnIndex(spreadsheet, sheet_title, column)

    if index < 0:
        return

    if location <= index:
        location -= 1
    request = SHEET_SERVICE.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['id'], body={
        "requests": [
            {
                "moveDimension": {
                    "source": {
                        "sheetId": Info.getSheetId(spreadsheet, sheet_title),
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
        Log.err(spreadsheet, err, True)