from openpyxl import *
from os import listdir
from os.path import isfile, join, exists
from tkinter import *
import re
import InputData as data
import xlwings as xw
import time
import math
from re import match
import copy
import traceback
import exception_logger as exlogger
import helper
from openpyxl.cell.cell import MergedCell

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
            'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

merge_sheet_prefix = helper.PREFIX_MERGE


def get_merge_cell_value(sheet, cell):
    rng = [s for s in sheet.merged_cells.ranges if cell.coordinate in s]
    return sheet.cell(rng[0].min_row, rng[0].min_col) if len(rng) != 0 else cell


def get_files(directory):
    directory = directory.replace("\\", "/")
    files = [f for f in listdir(directory) if isfile(
        join(directory, f)) and f.endswith(".xlsx")]
    files = [directory + f"/{f}" for f in files]
    return files


def remove_message_box(top_level, after_action=None):
    top_level.destroy()
    top_level.update()
    if after_action == None:
        return
    after_action()


def add_message_box(message, parent_frame, width=150, height=120, unmount_action=None):
    top = Toplevel(parent_frame, height=height, width=width)
    x = parent_frame.winfo_x()
    y = parent_frame.winfo_y()
    top.geometry("%dx%d+%d+%d" % (width, height, x + 200, y + 200))
    top.title("Info")
    Label(top, text=message, font=('Arial 11 normal'), pady=20).pack()
    Button(top, text="Ok", font=('Arial 11 normal'),
    command=lambda: remove_message_box(top, unmount_action)).pack()


def validate_directory(directory):
    if directory == "" or not exists(directory):
        return False
    return True


def validate_directory_and_select_sheet(directory, select_sheet, window):
    is_valid_directory = validate_directory(directory)
    if not is_valid_directory:
        add_message_box("Directory is invalid!", window)
        return False
    try:
        select_sheet = int(select_sheet)
        select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
    except:
        if select_sheet != data.INFINITE_SHEET_CHECKER:
            add_message_box("Invalid sheet number", window)
            exlogger.write(traceback.format_exc())
            return False
    return True


def change_data(command_list, directory, window, select_sheet):
    is_valid = validate_directory_and_select_sheet(
        directory, select_sheet, window)
    if not is_valid:
        return
    files = get_files(directory)
    for file in files:
        try:
            wb = load_workbook(filename=file)
            worksheets = []
            if select_sheet == data.INFINITE_SHEET_CHECKER:
                worksheets = wb.worksheets
            else:
                select_sheet = int(select_sheet)
                select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
                worksheets = [wb.worksheets[select_sheet]]
            for worksheet in worksheets:
                for command in command_list:
                    update_cell = command.cell
                    is_match = re.match(r"^[a-zA-Z]+\d+$", update_cell)
                    if not is_match:
                        if update_cell == "":
                            add_message_box("Empty cell value", window)
                            return
                        add_message_box(
                            f"'{update_cell}' isn't a cell", window)
                        return
                    if isinstance(worksheet[update_cell], MergedCell):
                        merge_update_cell = get_merge_cell_value(
                            worksheet, worksheet[update_cell])
                        merge_update_cell.value = command.value
                        continue
                    update_value = command.value
                    worksheet[update_cell] = update_value
            wb.save(file)
        except Exception as ex:
            # exception: index out of range
            if type(ex) == IndexError:
                file_name_log = file.split("/")[-1]
                add_message_box(
                    f"Error {file_name_log} \n out of range", window, height=140, width=150)
                return
            if type(ex) == PermissionError:
                add_message_box("Turn off all of\n .xlsx files first",
                                window, height=140, width=150)
                return
            # exception: all
            add_message_box("Error detected!", window)
            exlogger.write(traceback.format_exc())
            return
    add_message_box("Execute completed!", window)


def get_column_by_index(index):
    global alphabet
    divide_val = int(index/26)-1
    offset_val = int(index % 26)-1
    if divide_val == -1:
        return alphabet[offset_val]
    return f"{alphabet[divide_val]}{alphabet[offset_val]}"


def control_home(directory, select_sheet, window, process_bar_increase_val):
    is_valid = validate_directory_and_select_sheet(
        directory, select_sheet, window)
    if not is_valid:
        return
    files = get_files(directory)
    try:
        app = xw.App(visible=helper.OPEN_EXCEL_WHILE_RUNNING)
        steps = 100/len(files)

        for file in files:
            book = xw.Book(file)
            active_window = book.app.api.ActiveWindow
            for sheet in book.sheets:
                sheet.activate()
                is_freeze_panes = active_window.FreezePanes

                if is_freeze_panes:
                    freeze_row = int(active_window.SplitRow) + 1
                    freeze_column = int(active_window.SplitColumn) + 1
                    freeze_column = get_column_by_index(freeze_column)
                    select_coord = f"{freeze_column}{freeze_row}"
                    sheet.range(select_coord).select()
                else:
                    sheet.range("A1").select()
            book.sheets[0].activate()
            book.save(file)
            book.close()
            process_bar_increase_val(steps)
        add_message_box("Execute completed", window,
                        unmount_action=lambda: process_bar_increase_val(0))

    except Exception as ex:
        add_message_box("Error detected!", window)
        exlogger.write(traceback.format_exc())
    finally:
        app.quit()


def get_max_number(strings):
    max_number = None
    global merge_sheet_prefix
    pattern_regex = rf"\b{merge_sheet_prefix}(\d+)\b"

    for string in strings:
        match = re.search(pattern_regex, string)
        if match:
            number = int(match.group(1))
            if max_number is None or number > max_number:
                max_number = number

    if max_number == None:
        return 0
    return max_number


def check_if_merge_index_of_sheet(choose_sheet):
    try:
        select_sheet = int(choose_sheet)
        select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
        return select_sheet
    except Exception as ex:
        return "Not Set"


def delete_old_merge_sheet(destination_work_book, match_list):
    for match_item in match_list:
        destination_work_book.sheets[match_item].delete()


def merge_sheet(directory, choose_sheet, window, destination_file, is_remove_old_merge_sheet, process_bar_increase_val):
    try:
        app = xw.App(visible=helper.OPEN_EXCEL_WHILE_RUNNING)
        is_valid = validate_directory_and_select_sheet(
            directory, choose_sheet, window)
        if not is_valid:
            return
        files = get_files(directory)

        destination_file = destination_file.replace("\\", "/")

        if destination_file == "" or not isfile(destination_file):
            if destination_file == "":
                destination_file = directory.replace(
                    "\\", "/") + "/" + helper.DEFAULT_MERGE_FILE_NAME

            create_work_book = xw.Book()
            create_work_book.sheets[0].name = "Thống Kê"
            create_work_book.save(destination_file)
            create_work_book.close()

        destination_work_book = xw.Book(destination_file)
        destination_work_book_sheet_book_list = destination_work_book.sheet_names

        global merge_sheet_prefix
        match_list = list(filter(lambda v: match(
            rf'\b{merge_sheet_prefix}\d+\b', v), destination_work_book_sheet_book_list))

        if is_remove_old_merge_sheet:
            delete_old_merge_sheet(destination_work_book, match_list)
            match_list = []

        max_value = get_max_number(match_list)
        sheet_merge_name_index = max_value
        merge_index_sheet = check_if_merge_index_of_sheet(choose_sheet)
        steps = int(100/(sum(1 for i in files if i != destination_file)))
        for file in files:
            if file == destination_file:
                continue
            # sheet you want to copy
            source = xw.Book(file)
            sheet_list = source.sheet_names
            sheet_list_loop = None
            if merge_index_sheet == "Not Set":
                sheet_list_loop = sheet_list
            else:
                sheet_list_loop = [sheet_list[merge_index_sheet]]

            for sheet in sheet_list_loop:
                sheet_merge_name_index = sheet_merge_name_index + 1
                global merge_sheet_prefix
                merge_sheet_name = merge_sheet_prefix + \
                    str(sheet_merge_name_index)
                source.sheets[sheet].copy(name=merge_sheet_name, after=destination_work_book.sheets[len(
                    destination_work_book.sheets)-1])
            process_bar_increase_val(steps)

        destination_work_book.save()
        destination_work_book.close()
        add_message_box("Execute completed", window,
                        unmount_action=lambda: process_bar_increase_val(0))

    except Exception as ex:
        if isinstance(ex, PermissionError) or type(ex).__name__ == "com_error":
            add_message_box("Turn off all .xlsx files first ",
                            window, width=200)
            return
        print(ex)
        if isinstance(ex, KeyError):
            add_message_box("some file is broken", window, width=200)
            return
        exlogger.write(traceback.format_exc())
        add_message_box("Error Detect!", window)
        return
    finally:
        app.quit()


def translate_sheet(directory,select_language, choose_sheet, window, process_bar_increase_val):
    try:
        app = xw.App(visible=helper.OPEN_EXCEL_WHILE_RUNNING)
        is_valid = validate_directory_and_select_sheet(
            directory, choose_sheet, window)
        if not is_valid:
            return
        files = get_files(directory)
        print(select_language)
        print(files)

    except Exception as ex:
        print(ex)
