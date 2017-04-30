import json

with open('./template.json') as data_file:
    _TEMPLATE = json.load(data_file)


def update(func):
    def wrapper(*args, **kwargs):
        func(*args, **kwargs)
        with open('./template.json', 'w') as outfile:
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
    _TEMPLATE['sheets'].insert(index-1, {
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

def test(title):
    pass