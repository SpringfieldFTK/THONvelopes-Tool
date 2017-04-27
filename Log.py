import re


def err(spreadsheet=None, error="", raw=False):
    if raw:
        m = re.search("\".+?: (.+)\"", str(error))
        error = m.group(1)
    if spreadsheet:
        print("\nERROR in {0}({1}): {2}".format(spreadsheet['name'], spreadsheet['id'], error))
    else:
        print("\nERROR: {0}".format(error))