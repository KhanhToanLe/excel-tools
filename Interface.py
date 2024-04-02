import xlsxHelper as xlsx
from tkinter import *
import tkinter.filedialog
from tkinter import ttk
import helper as HELPER
import InputData as data

input_directory_entry = ""

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
	global input_directory_entry
	input_directory_entry =  input_directory
	button = Button(pane,text = "Select",background ="white", fg = "black",command=lambda:select_working_directory(input_directory))
	button.pack(side=LEFT,pady=4,padx=12);

def display_change_data_content_tab(tab1):
	container = Frame(tab1,background="white")
	height = HELPER.WINDOW_SIZE.split("x")[1]
	canvas = Canvas(container,background="white",height=(int(height)-100))
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
	add_change_data_click_handler(scrollable_frame,1)
	for i in data.change_data:
		i.frame.pack()
	return scrollable_frame

def reload_change_data_list(scrollable_frame):
	for i in scrollable_frame.winfo_children():
		i.pack_forget()
	for i in data.change_data:
		i.frame.pack(fill=X,expand=True)

def execute_change_data():
	current_frame = data.change_data[0].row.winfo_children()
	for change_value in current_frame:
		print(change_value.get())

def remove_change_data(scrollable_frame,added_frame):
	data.change_data.remove(added_frame)
	if len(data.change_data) == 0:
		add_change_data_click_handler(scrollable_frame,1)
	reload_change_data_list(scrollable_frame)

def upper_case_cell(event:Event):
	entry = event.widget
	primitive_val = entry.get()
	entry.delete(0,END)
	entry.insert(0,primitive_val.upper())

def add_change_data_click_handler(scrollable_frame,add_item):
	test = Frame(scrollable_frame)
	cell = Entry(test,font=("Arial", 11, "normal"),width=12 )
	cell.bind("<FocusOut>",upper_case_cell)
	cell.pack(side=LEFT)
	value = Entry(test,font=("Arial", 11, "normal"),width=37)
	value.pack(side=LEFT,fill='x', expand=True)
	added_frame = data.InputData(test)
	Button(test,width=4,text="+",command=lambda:add_change_data_click_handler(scrollable_frame,1)).pack(side=LEFT)
	Button(test,width=4,text="-",command=lambda:remove_change_data(scrollable_frame,added_frame)).pack(side=LEFT)
	if add_item == 1: 
		data.change_data.append(added_frame)
	else:
		insert_after = data.chang_data.index(add_item)
		data.change_data.insert(insert_after,added_frame)
	reload_change_data_list(scrollable_frame)



is_check_all_page = 0
old_choose_value = ""
chose_sheet = 1

def checker_all_page_handler(is_check):
	# global is_check_all_page
	global is_check_all_page, chose_sheet,old_choose_value
	is_check_all_page =  is_check.get()
	if(is_check_all_page == 1):
		old_choose_value = chose_sheet.get()
		chose_sheet.delete(0,END)
		chose_sheet.insert(0,'∞')
		chose_sheet.config(state= "disabled")
	else:
		chose_sheet.config(state= "normal")
		chose_sheet.delete(0,END)
		chose_sheet.insert(0,old_choose_value)

def display_change_data_execute_selection(window):
	execute_frame = Frame(window)
	execute_frame.pack(fill=BOTH,expand=True)

	Label(execute_frame, text="Select Sheet").grid(row=0,column=1)

	select_sheet = Entry(execute_frame,width=4)
	select_sheet.grid(row=0,column=2,padx=4)
	global chose_sheet
	chose_sheet = select_sheet
	# global is_check_all_page
	myVar = IntVar()
	Checkbutton(execute_frame, text='Every Sheet',variable=myVar, onvalue=1, offvalue=0, command=lambda:checker_all_page_handler(myVar)).grid(row=0,column=3)
	
	executeButton = Button(execute_frame, text ="Execute", command = lambda:change_data())
	executeButton.grid(row=0,column=4)

def display_control_home(control_home_tab):
	print("go here")

def	get_change_data_command():
	data.command_data = []
	for item in data.change_data:
		current_frame = item.frame
		entries = current_frame.winfo_children()
		data.command_data.append(data.ChangeData(entries[0].get(),entries[1].get()))

main_window = ""

def change_data():
	get_change_data_command()
	global chose_sheet
	xlsx.change_data(data.command_data,input_directory_entry.get(),main_window,chose_sheet.get())

def startup():
	window = create_main_window()
	global main_window
	main_window = window
	display_directory_input(window)
	change_data_tab, ctr_home_tab = display_tab(window)
	scrollable_frame = display_change_data_content_tab(change_data_tab)
	display_control_home(ctr_home_tab)
	display_change_data_execute_selection(window)
	return window, change_data_tab, ctr_home_tab,scrollable_frame



