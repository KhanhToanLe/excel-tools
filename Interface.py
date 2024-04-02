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
	pane.pack(fill = BOTH)
	Label(pane,text="Directory").pack(side=LEFT,pady=4,padx=4)
	border_color = Frame(pane, background="gray")
	input_directory = Entry(border_color)
	border_color.pack(padx=1,pady=1,fill='x',side=LEFT,expand=TRUE)
	input_directory.pack(padx=1,pady=1,fill='x')
	button = Button(pane,text = "Select",background = "white", fg = "black",command=lambda:select_working_directory(input_directory))
	button.pack(side=LEFT,pady=4,padx=12);

def display_change_data_content_tab(tab1):
	container = Frame(tab1,background="red")
	height = HELPER.WINDOW_SIZE.split("x")[1]
	canvas = Canvas(container,background="red",height=(int(height)-100))
	scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
	scrollable_frame = Frame(canvas)
	scrollable_frame.bind(
		"<Configure>",
		lambda e: canvas.configure(
			scrollregion=canvas.bbox("all")
		)
	)
	canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
	canvas.configure(yscrollcommand=scrollbar.set)
	container.pack(fill="both",)
	canvas.pack(side="left", fill=BOTH, expand=True)
	scrollbar.pack(side="right", fill="y")
	# button
	# AddButton = Button(container, text ="Add value", command = lambda:add_change_data_click_handler(scrollable_frame))
	# AddButton.place(x=4,y=4)
	# add_change_data_click_handler(scrollable_frame)
	add_change_data_click_handler(scrollable_frame)
	for i in data.change_data:
		i.cell.pack()
	return scrollable_frame

def reload_change_data_list():
	for i in data.change_data:
		i.cell.pack()

def execute_change_data():
	current_frame = data.change_data[0].row.winfo_children()
	for change_value in current_frame:
		print(change_value.get())

def add_change_data_click_handler(scrollable_frame):
	test = Frame(scrollable_frame)
	cell = Entry(test,font=("Arial", 11, "normal"),width=12 )
	cell.pack(side=LEFT)
	value = Entry(test,font=("Arial", 11, "normal"),width=12)
	value.pack(side=LEFT)
	Button(test,text="Add New",command=lambda:add_change_data_click_handler(scrollable_frame)).pack(side=LEFT)
	data.change_data.append(
		data.InputData(
			test
		)
	)
	reload_change_data_list()



def startup():
	window = create_main_window()
	display_directory_input(window)
	change_data_tab, ctr_home_tab = display_tab(window)
	scrollable_frame = display_change_data_content_tab(change_data_tab)
	return window, change_data_tab, ctr_home_tab,scrollable_frame



