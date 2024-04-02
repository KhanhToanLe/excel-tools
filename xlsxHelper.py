from openpyxl import *
from os import listdir
from os.path import isfile, join, exists
from tkinter import *
import re
import InputData as data	


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
	
def change_data(command_list, directory,window, select_sheet):
	is_valid_directory = validate_directory(directory)
	if not is_valid_directory :
		add_message_box("Directory is invalid!",window)
		return
	try: 
		select_sheet = int(select_sheet)
		select_sheet = 0 if select_sheet - 1 <= 0 else select_sheet - 1
	except:
		if select_sheet != data.INFINITE_SHEET_CHECKER:
			add_message_box("Invalid sheet number",window)
			return
	files = get_files(directory)
	for file in files:
		try:
			wb = load_workbook(filename = file)
			worksheets = []
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
					
			# worksheet.sheet_view.topLeftCell = 'A1'
			
			worksheet.views.sheetView[0].selection[0].activeCell = "topLeft"
			worksheet.views.sheetView[0].selection[0].sqref = 'topLeft'
			print(worksheet.freeze_panes)
			wb.save(file)
		except Exception as ex:
			# exception: index out of range
			if type(ex) ==  IndexError:
				file_name_log = file.split("/")[-1]
				add_message_box(f"Error {file_name_log} \n out of range",window,height=140,width=150)
				return
			if type(ex) == AttributeError:
				add_message_box("Turn off .xlsx files",window,height=140,width=150)
				return
			# exception: all
			print(ex)
			add_message_box("Error detect!",window)
			return
	# complete
	add_message_box("Execute completed!",window)

