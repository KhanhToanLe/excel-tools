import bussiness as bsn
from tkinter import *
import tkinter.filedialog
from tkinter import ttk
import helper as HELPER
import InputData as data

def create_main_window():
	main_window = Tk()
	# main window set up
	# main_window.state('zoomed')
	main_window.geometry(HELPER.WINDOW_SIZE)
	main_window.title(HELPER.APP_NAME)

	return main_window

def select_working_directory(input_directory:Entry):
	select_directory = tkinter.filedialog.askdirectory()
	input_directory.delete(0,END)
	input_directory.insert(0,select_directory)

def display_tab(window):
	tabControl = ttk.Notebook(window) 
	tab1 = Frame(tabControl)
	tab2 = Frame(tabControl)
	tabControl.add(tab1, text ='Change Data') 
	tabControl.add(tab2, text ='Ctrl Home') 
	tabControl.pack(expand = 1, fill ="both")
	return tab1,tab2

def display_directory_input(window):
	pane = Frame(window)
	pane.pack(fill = 'x')
	Label(pane,text="Directory").pack(side=LEFT,pady=4,padx=4)
	border_color = Frame(pane, background="gray")
	input_directory = Entry(border_color)
	border_color.pack(padx=1,pady=1,fill='x',side=LEFT,expand=TRUE)
	input_directory.pack(padx=1,pady=1,fill='x')
	button = Button(pane,text = "Select",background = "white", fg = "black",command=lambda:select_working_directory(input_directory))
	button.pack(side=LEFT,pady=4,padx=12);

def display_change_data_content_tab():
	for change_value in data.change_data:
		change_value.row.pack();
		# change_value.column.pack();
		# change_value.value.pack();

def execute_change_data():
	current_frame = data.change_data[0].row.winfo_children()
	for change_value in current_frame:
		print(change_value.get())

def add_change_data_click_handler(window):
	pane = Frame(window);
	Entry(pane, width=8,relief="solid",font=('Arial', 11, 'normal')).pack(padx=4,pady=4,side=LEFT,expand=TRUE)
	Entry(pane, width=8,relief="solid",font=('Arial', 11, 'normal')).pack(padx=4,pady=4,side=LEFT,expand=TRUE)
	Entry(pane, width=8,relief="solid",font=('Arial', 11, 'normal')).pack(padx=4,pady=4,side=LEFT,expand=TRUE)

	data.change_data.append(
		data.InputData(
			pane
		)
	)
	display_change_data_content_tab()

def startup():
	window = create_main_window()
	display_directory_input(window)
	change_data_tab, ctr_home_tab = display_tab(window)
	display_change_data_content_tab()
	return window, change_data_tab, ctr_home_tab



