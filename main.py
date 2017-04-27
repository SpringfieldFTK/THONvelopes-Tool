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

#thonvelopes.main()
#thonvelopes.addSheet("Bulk Sheet")

visuals.create_gui(action)

#Bulk.deleteSheet("Bulk Sheet")
