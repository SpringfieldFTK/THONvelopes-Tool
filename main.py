import threading

from appJar import gui
import visuals
import Info
import Template

from Requests import Bulk, Single


def action(btn, extra=None):
    print(btn)
    if btn == "create_single_file":
        name = visuals.get_app().getEntry('name')
        if not name:
            visuals.set_info_error("Name cannot be empty")
            return
        Single.createFile(name)
        visuals.set_info_success("File {0} successfully created".format(name))
        visuals.get_app().setEntry('name', "")
    elif btn == "create_multiple_files":
        names = visuals.get_app().getTextArea('names')
        if not names:
            visuals.set_info_error("Names cannot be empty")
            return
        suffix = visuals.get_app().getEntry('suffix')
        names = names.split('\n')

        Single.batchCreateFile(names, suffix)

        visuals.set_info_success("{0} files successfully created".format(len(names)))
        visuals.get_app().clearTextArea('names')
        visuals.get_app().setEntry('suffix', "")

    elif btn == "delete_sheet":
        title = visuals.get_app().getTabbedFrameSelectedTab("SheetFrame")
        Bulk.deleteSheet(title)
        Template.delete_sheet(title)
        Info.deleteSheet(title)
        visuals.edit_files()

    elif btn == "add_sheet":
        title = visuals.get_app().getEntry("title")
        cols = visuals.get_app().getTextArea("cols").split('\n')
        index = int(visuals.get_app().getSpinBox("Location in file"))
        Bulk.addSheet(title, cols, index)
        Template.add_sheet(title, cols, index)
        visuals.get_app().setEntry("title", "")
        visuals.get_app().clearTextArea("cols")

#thonvelopes.main()
#thonvelopes.addSheet("Bulk Sheet")

visuals.create_gui(action)

#Bulk.deleteSheet("Bulk Sheet")
