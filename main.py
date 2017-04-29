import threading

from appJar import gui
import visuals

from Requests import Bulk, Single



def action(btn):
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

#thonvelopes.main()
#thonvelopes.addSheet("Bulk Sheet")

visuals.create_gui(action)

#Bulk.deleteSheet("Bulk Sheet")
