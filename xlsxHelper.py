from openpyxl import *
from os import listdir
from os.path import isfile, join, exists
from tkinter import *
import re
import InputData as data
import xlwings as xw
import time


alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

def get_files(directory):
	directory = directory.replace("\\","/")
	files = [f for f in listdir(directory) if isfile(join(directory, f)) and f.endswith(".xlsx")]
	files = [directory + f"/{f}" for f in files]
	return files

def remove_message_box(top_level):
	top_level.destroy()
	top_level.update()

def add_message_box(message,parent_frame,width=150,height=120):
	top= Toplevel(parent_frame,height=height,width=width)
	x = parent_frame.winfo_x()
	y = parent_frame.winfo_y()
	top.geometry("%dx%d+%d+%d" % (width, height, x + 200, y + 200))
	top.title("Info")
	Label(top, text= message, font=('Arial 11 normal'),pady=20).pack()
	Button(top, text= "Ok", font=('Arial 11 normal'),command=lambda:remove_message_box(top)).pack()

def validate_directory(directory):
	if directory == "" or not exists(directory):
		return False
	return True

def validate_directory_and_select_sheet(directory, select_sheet, window):
	is_valid_directory = validate_directory(directory)
	if not is_valid_directory :
		add_message_box("Directory is invalid!",window)
		return False
	try: 
		select_sheet = int(select_sheet)
		select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
	except:
		if select_sheet != data.INFINITE_SHEET_CHECKER:
			add_message_box("Invalid sheet number",window)
			return False
	return True

	
def change_data(command_list, directory,window, select_sheet):
	is_valid = validate_directory_and_select_sheet(directory,select_sheet,window)
	if not is_valid: 
		return
	files = get_files(directory)
	for file in files:
		try:
			wb = load_workbook(filename = file)
			worksheets = []
			select_sheet = int(select_sheet)
			select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
			if select_sheet != data.INFINITE_SHEET_CHECKER:
				worksheets = [wb.worksheets[select_sheet]]
			else:
				worksheets = wb.worksheets
			for worksheet in worksheets:
				for command in command_list: 
					update_cell = command.cell
					is_match = re.match(r"^[a-zA-Z]+\d+$",update_cell)
					if not is_match:
						if update_cell == "":
							add_message_box("Empty cell value",window)
							return
						add_message_box(f"'{update_cell}' isn't a cell",window)
						return
					update_value = command.value
					worksheet[update_cell] = update_value
			wb.save(file)
		except Exception as ex:
			# exception: index out of range
			if type(ex) ==  IndexError:
				file_name_log = file.split("/")[-1]
				add_message_box(f"Error {file_name_log} \n out of range",window,height=140,width=150)
				return
			if type(ex) == PermissionError:
				add_message_box("Turn off all of\n .xlsx files first",window,height=140,width=150)
				return
			# exception: all
			print(type(ex))
			add_message_box("Error detected!",window)
			return
	# complete
	add_message_box("Execute completed!",window)
def get_column_by_index(index):
	global alphabet
	divide_val = int(index/26)-1
	offset_val = int(index%26)-1
	if divide_val == -1:
		return alphabet[offset_val]
	return f"{alphabet[divide_val]}{alphabet[offset_val]}"

def control_home(directory, select_sheet, window,process_bar_increase_val):
	is_valid = validate_directory_and_select_sheet(directory,select_sheet, window)
	if not is_valid: 
		return
	files = get_files(directory)
	num_files = len(files)
	num_steps =  100/num_files
	try:
		app = xw.App(visible=True)
		process_bar_increase_val(num_steps)
		for file in files:
			book = xw.Book(file)
			active_window = book.app.api.ActiveWindow
			for sheet in book.sheets:
				sheet.activate()
				is_freeze_panes = active_window.FreezePanes
				if is_freeze_panes:
					freeze_row =  int(active_window.SplitRow) + 1
					freeze_column = int(active_window.SplitColumn) + 1
					freeze_column = get_column_by_index(freeze_column)
					select_coord = f"{freeze_column}{freeze_row}"
					sheet.range(select_coord).select()
				else:
					sheet.range("A1").select()
			# TODO: jump to sheet feature
			book.save(file)
			book.close()
			time.sleep(0.5)

		add_message_box("Execute completed",window)
		app.quit()
	except Exception as ex:
		print(ex)
		add_message_box("Error detected!",window)



