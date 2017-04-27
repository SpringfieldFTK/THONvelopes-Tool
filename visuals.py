from appJar import gui


def up_one(up_function):
    def real_decorator(function):
        def wrapper(*args, **kwargs):
            global row
            app.setSticky("new")
            app.setExpand("column")
            app.addButton("Back", up_function, row, 0)
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
    app.setMessageFg('info','green')
    app.showMessage("info")


def set_info_error(message):
    app.setMessage("info", message)
    app.setMessageFg('info','red')
    app.showMessage("info")


def reset(visual_function):
    def wrapper(*args, **kwargs):
        global row
        row = 0
        app.removeAllWidgets()
        visual_function(*args, **kwargs)
    return wrapper


@reset
def main_menu(btn=None):
    global row
    app.removeAllWidgets()
    app.setSticky("news")
    app.setExpand("both")
    app.setFont(20)
    app.setPadding(5, 5)

    app.addNamedButton("Create\nNew File", "create", create_file_menu, row, 0)
    app.addNamedButton("Edit\nAll Files", "edit", create_file_menu, row, 1)
    row += 1
    app.addNamedButton("Consolidate\nData", "consolidate", create_file_menu, row, 0)
    app.addNamedButton("Disperse\nData", "disperse", create_file_menu, row, 1)

    app.setButtonBg("create", "LightYellow")
    app.setButtonBg("edit", "LemonChiffon")
    app.setButtonBg("consolidate", "PapayaWhip")
    app.setButtonBg("disperse", "Moccasin")


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
def single_file(btn = None):
    global row
    app.addLabel("desc", "Enter the name for the file below:", row, 0, 2)
    row += 1
    app.addEntry("name", row, 0, 2)
    app.setEntryDefault("name", "File Name")
    row += 1
    app.addNamedButton("Create", "create_single_file",  action, row, 0, 2)


@reset
@up_one(create_file_menu)
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


def create_gui(action_method):
    global app, action
    action = action_method
    main_menu()
    app.go()


def get_app():
    global app
    return app

app = gui("THONvelopes", "500x500")
action = None
row = 0

