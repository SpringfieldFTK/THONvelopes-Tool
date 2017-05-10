import json

_TEMPLATE = None


def update(func):
    def wrapper(*args, **kwargs):
        func(*args, **kwargs)
        with open('./Resources/template.json', 'w') as outfile:
            json.dump(_TEMPLATE, outfile)

    return wrapper


def get_sheets():
    return [x['title'] for x in _TEMPLATE['sheets']]


def get_header_row(sheet):
    for s in _TEMPLATE['sheets']:
        if s['title'] == sheet:
            return s['header_row']
    return None


def get_columns(sheet):
    for s in _TEMPLATE['sheets']:
        if s['title'] == sheet:
            return s['header_columns']
    return None


@update
def add_sheet(title, cols, index):
    if "sheets" not in _TEMPLATE:
        _TEMPLATE["sheets"] = []
    _TEMPLATE['sheets'].insert(index - 1, {
        "title": title,
        "header_row": 1,
        "header_columns": cols
    })


@update
def delete_sheet(title):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            _TEMPLATE['sheets'].pop(index)
            break


@update
def rename_sheet(title, new_title):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            _TEMPLATE['sheets'][index]['title'] = new_title
            break


@update
def moveSheet(title, new_index):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            moving_sheet = _TEMPLATE['sheets'].pop(index)
            if new_index > index:
                new_index += 1
            _TEMPLATE['sheets'].insert(new_index, moving_sheet)
            break
    return None


@update
def delete_column(title, col):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            _TEMPLATE['sheets'][index]['header_columns'].remove(col)
            break


@update
def rename_column(title, col, new_col):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            index = _TEMPLATE['sheets'][index]['header_columns'].index(col)
            _TEMPLATE['sheets'][index]['header_columns'][index] = new_col
            break


@update
def add_column(title, new_col):
    for index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            _TEMPLATE['sheets'][index]['header_columns'].append(new_col)
            break


@update
def set_template(js):
    global _TEMPLATE
    _TEMPLATE = js


@update
def move_column(title, col, new_index):
    for sheet_index, sheet in enumerate(_TEMPLATE['sheets']):
        if sheet['title'] == title:
            print(_TEMPLATE['sheets'][sheet_index]['header_columns'])
            old_index = _TEMPLATE['sheets'][sheet_index]['header_columns'].index(col)
            print(old_index)
            moving_col = _TEMPLATE['sheets'][sheet_index]['header_columns'].pop(old_index)
            print(_TEMPLATE['sheets'][sheet_index]['header_columns'])
            print(new_index)
            _TEMPLATE['sheets'][sheet_index]['header_columns'].insert(new_index, moving_col)
            print(_TEMPLATE['sheets'][sheet_index]['header_columns'])
            break
    return None
