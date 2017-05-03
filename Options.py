import json
import os.path
import visuals

if os.path.isfile("./Resources/options.json"):
    with open('./Resources/options.json') as data_file:
        _OPTIONS = json.load(data_file)
else:
    _OPTIONS = {}
    with open('./Resources/options.json', 'w') as outfile:
        json.dump(_OPTIONS, outfile)


def update(func):
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        with open('./Resources/options.json', 'w') as outfile:
            json.dump(_OPTIONS, outfile)
        return result
    return wrapper


@update
def get_indiv_folder():
    if "individual_files" in _OPTIONS:
        return _OPTIONS["individual_files"]
    else:
        id = visuals.get_text_input("Folder ID", "Please paste the ID of the folder containing individual THONvelope files")
        if not id:
            exit()
        _OPTIONS["individual_files"] = id
        return id
