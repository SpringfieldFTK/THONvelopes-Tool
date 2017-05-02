import pprint

from appJar import gui

import Template


def up_one(up_function):
    def real_decorator(function):
        def wrapper(*args, **kwargs):
            global row
            app.setSticky("new")
            app.setExpand("column")
            app.addButton("Back", up_function, row, 0)
            app.setButtonBg("Back", "LightSteelBlue")
            row += 1
            function(*args, **kwargs)

        return wrapper

    return real_decorator


def info(info_function):
    def wrapper(*args, **kwargs):
        global row
        app.setSticky("new")
        app.setExpand("column")
        app.addEmptyMessage("info", row, 0, 2)
        row += 1
        app.hideMessage('info')
        info_function(*args, **kwargs)

    return wrapper


def set_info_success(message):
    app.setMessage("info", message)
    app.setMessageFg('info', 'green')
    app.showMessage("info")


def set_info_error(message):
    app.setMessage("info", message)
    app.setMessageFg('info', 'red')
    app.showMessage("info")


def reset(visual_function):
    def wrapper(*args, **kwargs):
        global row
        row = 0
        app.removeAllWidgets()
        visual_function(*args, **kwargs)

    return wrapper


def progress(function):
    def wrapper(*args, **kwargs):
        global row
        app.addMeter("prog", row, 0, 2)
        app.setMeterFill("prog", "blue")
        app.hideMeter("prog")
        row += 1
        function(*args, **kwargs)

    return wrapper


def set_progress(completed, max):
    app.showMeter("prog")
    app.setMeter('prog', float(completed) / max * 100)


@reset
def main_menu(btn=None):
    global row
    app.removeAllWidgets()
    app.setSticky("news")
    app.setExpand("both")
    app.setFont(20)
    app.setPadding(5, 5)

    app.addNamedButton("Create\nNew File", "create", create_file_menu, row, 0)
    app.addNamedButton("Edit\nAll Files", "edit", edit_files, row, 1)
    row += 1
    app.addNamedButton("Consolidate\nData", "consolidate", create_file_menu, row, 0)
    app.addNamedButton("Disperse\nData", "disperse", create_file_menu, row, 1)

    app.setButtonBg("create", "LightYellow")
    app.setButtonBg("edit", "LemonChiffon")
    app.setButtonBg("consolidate", "PapayaWhip")
    app.setButtonBg("disperse", "Moccasin")

    app.disableButton("consolidate")
    app.disableButton("disperse")


@reset
@up_one(main_menu)
def create_file_menu(btn=None):
    app.setSticky("news")
    app.setExpand("both")
    app.setFont(20)
    app.setPadding(5, 5)

    app.addNamedButton("Single File", "single", single_file, row, 0, 1, 2)
    app.addNamedButton("Multiple Files", "batch", multiple_files, row, 1, 1, 2)

    app.setButtonBg("single", "LightYellow")
    app.setButtonBg("batch", "LemonChiffon")


@reset
@up_one(create_file_menu)
@info
def single_file(btn=None):
    global row
    app.addLabel("desc", "Enter the name for the file below:", row, 0, 2)
    row += 1
    app.addEntry("name", row, 0, 2)
    app.setEntryDefault("name", "File Name")
    row += 1
    app.addNamedButton("Create", "create_single_file", action, row, 0, 2)


@reset
@up_one(create_file_menu)
@progress
@info
def multiple_files(btn=None):
    global row
    app.setFont(14)

    app.addLabel("desc1", "Enter one member name per line below:", row, 0, 2)
    row += 1

    app.addScrolledTextArea("names", row, 0, 2)
    row += 1

    app.addLabel("desc2", "Enter a suffix to put after all names below (optional):", row, 0, 2)
    row += 1

    app.addEntry("suffix", row, 0, 2)
    row += 1

    app.addNamedButton("Create All", "create_multiple_files", action, row, 0, 2)
    row += 1

    app.setTextAreaHeight("names", 10)


@reset
@up_one(main_menu)
@progress
def edit_files(btn=None):
    global row, action
    app.addNamedButton("Add Sheet", "add_sheet", add_sheet, row, 0, 2, 0)
    app.setButtonBg("add_sheet", "LightGreen")
    row += 1
    app.startTabbedFrame("SheetFrame", row, 0, 2, 2, 5)
    for index, title in enumerate(Template.get_sheets()):
        app.startTab(title)
        app.setSticky("news")
        app.setExpand("both")
        app.setPadding(5, 5)
        app.addNamedButton("Edit Columns", title + "_columns", edit_cols, 0, 0, 3, 2)
        app.addNamedButton("Edit Rows", title + "_rows", main_menu, 0, 3, 3, 2)
        app.addNamedButton("Rename", title + "_rename", lambda x: action("rename_sheet"), 3, 0, 2, 1)
        app.addNamedButton("Move", title + "_move", lambda x: action("move_sheet"), 3, 2, 2, 1)
        app.addNamedButton("Delete", title + "_delete", lambda x: action("delete_sheet"), 3, 4, 2, 1)
        app.setButtonBg(title + "_delete", "Red")
        app.stopTab()

    app.stopTabbedFrame()


@reset
@up_one(edit_files)
@progress
def add_sheet(btn=None):
    global row
    app.setFont(14)
    app.addLabel("title_desc","Enter a title for the new sheet below:", row, 0, 2)
    row += 1
    app.addEntry("title",row, 0, 2)
    row += 1
    app.addLabel("cols_desc", "Enter the columns you would like in the new sheet below:", row, 0, 2)
    row += 1
    app.addScrolledTextArea("cols", row, 0, 2)
    row += 1
    app.addLabelSpinBoxRange("Location in file", 1, len(Template.get_sheets()) + 1, row, 1, 2)
    app.setSpinBox("Location in file", len(Template.get_sheets()) + 1, callFunction=False)
    row += 1
    app.addNamedButton("Add Sheet", "add_sheet", action, row, 0, 2)
    app.setTextAreaHeight("cols", 10)


@reset
@up_one(edit_files)
@progress
def edit_cols(btn):
    global row
    app.startLabelFrame("Columns", row, 0, 2)

    app.setSticky("news")
    app.setExpand("both")
    app.setPadding(5, 5)

    title = btn.split("_")[0]
    app.addLabel("title", "Editing columns in \n'{0}'".format(title),0,0,6)
    app.addNamedButton("Add Column", "add_col", action, 1, 0, 6)
    app.addOptionBox("cols", Template.get_columns(title), 2, 0, 6)
    app.addNamedButton("Rename", "rename_col", action, 3, 0, 2)
    app.addNamedButton("Move", "move_col", action, 3, 2, 2)
    app.addNamedButton("Delete", "delete_col", action, 3, 4, 2)

    app.stopLabelFrame()


def create_gui(action_method):
    global app, action
    action = action_method
    main_menu()
    app.setBg("Azure")
    app.go()


def get_app():
    global app
    return app


app = gui("THONvelopes", "500x500")
progress = gui("Progress", "500x200")
action = None
row = 0
max_prog = 0
current_prog = 0
