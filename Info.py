import string
import re

from googleapiclient.errors import HttpError

import Log
from Requests import SHEET_SERVICE

INFO = {}


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


def getSheetId(spreadsheet, title):
    if spreadsheet['id'] in INFO:
        if title in INFO[spreadsheet["id"]]:
            if "id" in INFO[spreadsheet["id"]][title]:
                return INFO[spreadsheet["id"]][title]["id"]
        else:
            INFO[spreadsheet["id"]][title] = {}
    else:
        INFO[spreadsheet["id"]] = {}
        INFO[spreadsheet["id"]][title] = {}

    request = SHEET_SERVICE.spreadsheets().get(spreadsheetId=spreadsheet['id'])

    try:
        response = request.execute()
        for sheet in response['sheets']:
            if sheet['properties']['title'].lower() == title.lower():
                id = sheet['properties']['sheetId']
                INFO[spreadsheet["id"]][title]["id"] = id
                return sheet['properties']['sheetId']
    except HttpError as err:
        Log.err(spreadsheet, err, True)
    Log.err(spreadsheet, "Sheet '{0}' not found".format(title))
    return -1


def getSheetIndex(spreadsheet, title):
    if spreadsheet['id'] in INFO:
        if title in INFO[spreadsheet["id"]]:
            if "index" in INFO[spreadsheet["id"]][title]:
                return INFO[spreadsheet["id"]][title]["index"]
        else:
            INFO[spreadsheet["id"]][title] = {}
    else:
        INFO[spreadsheet["id"]] = {}
        INFO[spreadsheet["id"]][title] = {}

    request = SHEET_SERVICE.spreadsheets().get(spreadsheetId=spreadsheet['id']).execute()
    try:
        response = request.execute()
        for sheet in response['sheets']:
            if sheet['properties']['title'].lower() == title.lower():
                index = sheet['properties']['index']
                INFO[spreadsheet["id"]][title]["index"] = index
                return index
    except HttpError as err:
        Log.err(spreadsheet, err, True)
    Log.err(spreadsheet, "Sheet '{0}' not found".format(title))
    return -1


def getColumnIndex(spreadsheet, sheet_title, contents):
    if spreadsheet['id'] in INFO:
        if sheet_title in INFO[spreadsheet["id"]]:
            if "columns" in INFO[spreadsheet["id"]][sheet_title]:
                if contents in INFO[spreadsheet["id"]][sheet_title]["columns"]:
                    return INFO[spreadsheet["id"]][sheet_title]["columns"][contents]
            else:
                INFO[spreadsheet["id"]][sheet_title]["columns"] = {}
        else:
            INFO[spreadsheet["id"]][sheet_title] = {}
            INFO[spreadsheet["id"]][sheet_title]["columns"] = {}
    else:
        INFO[spreadsheet["id"]] = {}
        INFO[spreadsheet["id"]][sheet_title] = {}
        INFO[spreadsheet["id"]][sheet_title]["columns"] = {}

    request = SHEET_SERVICE.spreadsheets().values().get(spreadsheetId=spreadsheet['id'],
                                                        range="'{0}'!A1:ZZZ1".format(sheet_title),
                                                        majorDimension="ROWS")
    try:
        response = request.execute()
        for i, cell in enumerate(response['values'][0]):
            if cell.lower() == contents.lower():
                INFO[spreadsheet["id"]][sheet_title]["columns"][contents] = i
                return i
    except HttpError as err:
        Log.err(spreadsheet, err, True)

    Log.err(spreadsheet, "Column '{0}' not found in '{1}'".format(contents, sheet_title))
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
        Log.err(spreadsheet, err, True)


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
        Log.err(spreadsheet, err, True)

    return -1