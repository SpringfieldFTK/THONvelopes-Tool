import json
import os
import threading

from appJar import gui
import visuals
import Info
import Template

from Requests import Bulk, Single


def action(btn):
    (globals()[btn])()

def create_single_file():
    name = visuals.get_app().getEntry('name')
    if not name:
        visuals.set_info_error("Name cannot be empty")
        return
    Single.createFile(name)
    visuals.set_info_success("File {0} successfully created".format(name))
    visuals.get_app().setEntry('name', "")


def create_multiple_files():
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


def delete_sheet():
    title = visuals.get_app().getTabbedFrameSelectedTab("SheetFrame")
    if visuals.get_app().yesNoBox("Confirm Delete", "Are you sure you delete sheet '{0}'?".format(title)):
        Bulk.deleteSheet(title)
        Template.delete_sheet(title)
        Info.deleteSheet(title)
        visuals.edit_files()


def add_sheet():
    title = visuals.get_app().getEntry("title")
    cols = visuals.get_app().getTextArea("cols").split('\n')
    index = int(visuals.get_app().getSpinBox("Location in file"))
    Bulk.addSheet(title, cols, index)
    Template.add_sheet(title, cols, index)
    visuals.get_app().setEntry("title", "")
    visuals.get_app().clearTextArea("cols")


def rename_sheet():
    title = visuals.get_app().getTabbedFrameSelectedTab("SheetFrame")
    new_title = visuals.get_app().textBox("Rename " + title, "Please enter a new name for sheet '{0}':".format(title))
    if new_title:
        Bulk.renameSheet(title, new_title)
        Template.rename_sheet(title, new_title)
        Info.rename_sheet(title, new_title)
        visuals.edit_files()


def move_sheet():
    title = visuals.get_app().getTabbedFrameSelectedTab("SheetFrame")
    max = len(Template.get_sheets())
    index = visuals.get_app().numberBox("New Index", "Please pick a new index, min of 1, max of {0}:".format(max))
    if index and 0 < index <= max:
        Bulk.moveSheet(title, index)
        Template.moveSheet(title, index)
        Info.invalidate_indexes()
        visuals.edit_files()
    elif index and (index < 0 or index > max):
        visuals.get_app().errorBox("Number Error", "The index you entered was outside the acceptable range")


def delete_col():
    sheet = visuals.get_app().getLabel("title").split("'")[1]
    col = visuals.get_app().getOptionBox("cols")
    if visuals.get_app().yesNoBox("Confirm Delete", "Are you sure you delete column '{0}' from sheet '{1}'?".format(col, sheet)):
        Bulk.delete_column(sheet, col)
        Template.delete_column(sheet, col)
        Info.delete_columns(sheet)
        visuals.get_app().changeOptionBox("cols", Template.get_columns(sheet), 0)


def rename_col():
    sheet = visuals.get_app().getLabel("title").split("'")[1]
    col = visuals.get_app().getOptionBox("cols")
    new_col = visuals.get_app().textBox("Rename Column", "Please enter a new name for column '{0}' in sheet '{1}':".format(col, sheet))
    if new_col:
        Bulk.rename_column(sheet, col, new_col)
        Template.rename_column(sheet, col, new_col)
        Info.rename_column(sheet, col, new_col)
        visuals.get_app().changeOptionBox("cols", Template.get_columns(sheet), 0)


def add_col():
    sheet = visuals.get_app().getLabel("title").split("'")[1]
    new_col = visuals.get_app().textBox("Add Column",
                                        "Please enter a name for the new columns in sheet '{0}':".format(sheet))
    if new_col:
        Bulk.add_column(sheet, new_col)
        Template.add_column(sheet, new_col)
        visuals.get_app().changeOptionBox("cols", Template.get_columns(sheet), 0)


def move_col():
    sheet = visuals.get_app().getLabel("title").split("'")[1]
    col = visuals.get_app().getOptionBox("cols")
    loc = visuals.get_app().textBox("Move Column",
                                        "Please enter the location (A, B, etc) to which you would like to move column "
                                        "'{0}' in sheet '{1}':".format(col, sheet))
    if loc:
        index = Info.col_to_index(loc)
        Bulk.move_column(sheet, col, index)
        Template.move_column(sheet, col, index)
        Info.delete_columns(sheet)
        visuals.get_app().changeOptionBox("cols", Template.get_columns(sheet), 0)

#thonvelopes.main()
#thonvelopes.addSheet("Bulk Sheet")


def import_template():
    if os.path.isfile("./Resources/template.json"):
        with open('./Resources/template.json') as data_file:
            temp = json.load(data_file)
    else:
        temp = {}
        id = visuals.get_text_input("Template ID", "Please paste the ID of a sample file")
        if not id:
            exit()
        sheets = Info.get_sheets(id)
        temp['sheets'] = []
        for title in sheets:
            row_num = visuals.get_number_input("Header Row", "What is the header row for sheet '{0}'".format(title))
            row = Info.get_row(id, title, row_num)
            temp['sheets'].append({
                'title': title,
                'header_row': row_num,
                'header_columns': row
            })
    Template.set_template(temp)

import_template()
visuals.create_gui(action)

#Bulk.deleteSheet("Bulk Sheet")
