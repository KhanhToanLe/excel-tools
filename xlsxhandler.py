from openpyxl import *
from os import listdir
from os.path import isfile, join

def get_all_xlsx_file(path):
    onlyfiles = [f for f in listdir(path) if isfile(join(path, f)) and f.endswith(".xlsx")]
    return onlyfiles

def change_data_a_file(file_path, change_list):
	wb = load_workbook(filename = file_path)
	active_sheets = wb.active
	input_cell = 'A1';
	active_sheets[input_cell] = "Hello World"
	wb.save(file_path)

def change_data(change_list, select_directory_entry):
	directory = select_directory_entry.get()
	directory = directory.replace("\\","/")
	files = get_all_xlsx_file(directory)
	files = [directory + f"/{x}" for x in files]
	for file in files:
		change_data_a_file(file,change_list)
	
# TODO: set active sheets =>>>
