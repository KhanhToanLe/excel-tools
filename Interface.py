import xlsxhandler as xlsxHdler
from tkinter import *
import tkinter.filedialog
from tkinter import ttk
import helper as HELPER
import InputData as data

selected_directory = ""

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
	global selected_directory
	selected_directory = input_directory
	button = Button(pane,text = "Select",background = "white", fg = "black",command=lambda:select_working_directory(input_directory))
	button.pack(side=LEFT,pady=4,padx=12);


def display_change_data_content_tab(tab1):
	container = Frame(tab1,background="red")
	height = HELPER.WINDOW_SIZE.split("x")[1]
	canvas = Canvas(container,background="white",height=(int(height)-90))
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

# optimize this
def reload_change_data_list(scrollable_frame):
	for frame in scrollable_frame.winfo_children():
		frame.pack_forget()
	for change_data in data.change_data:
		change_data.frame.pack()

def execute_change_data():
	for current_frame in data.change_data:
		current_value = current_frame.frame.winfo_children()
		# for change_value in current_value:
			# if isinstance(change_value, Entry):
			# print(change_value.get())
		data.excel_change_data_list.append(
			data.ExcelChangeData(current_value[0].get(),current_value[1].get()),
		)
	global selected_directory
	xlsxHdler.change_data(data.excel_change_data_list,selected_directory)


def remove_change_data_click_handler(scrollable_frame, remove_data):
	data.change_data.remove(remove_data)
	if(len(data.change_data) == 0):
		add_change_data_click_handler(scrollable_frame,1)
		return
	reload_change_data_list(scrollable_frame)

def add_change_data_click_handler(scrollable_frame,add_item):
	test = Frame(scrollable_frame)
	cell = Entry(test,font=("Arial", 11, "normal"),width=12 )
	cell.pack(side=LEFT)
	value = Entry(test,font=("Arial", 11, "normal"),width=12)
	value.pack(side=LEFT)
	add_change_item =  data.InputData(test)
	Button(test,text="+",command=lambda:add_change_data_click_handler(scrollable_frame,add_change_item),padx=6,pady=2).pack(side=LEFT)
	Button(test,text="-",command=lambda:remove_change_data_click_handler(scrollable_frame, add_change_item),padx=6,pady=2).pack(side=LEFT)
	if add_item == 1:
		data.change_data.append(add_change_item)
	else :
		insert_after = data.change_data.index(add_item)
		data.change_data.insert(insert_after + 1,add_change_item)
	reload_change_data_list(scrollable_frame)

def print_selection(value):
  print(value.get())

def display_execute_content_tab(window):
	execute_frame = Frame(window)
	execute_frame.pack(side=BOTTOM,fill=BOTH,expand=True)

	Label(execute_frame,text="Select sheet:").grid(row=0, column=1, sticky="e",pady=2)

	selected_page = Entry(execute_frame,width=8)
	selected_page.delete(0,END)
	selected_page.insert(0,1)
	selected_page.grid(row=0, column=2, sticky="n",pady=2)

	is_all_sheet = IntVar()
	is_all_sheet_checker = Checkbutton(execute_frame, text='Every sheet',variable=is_all_sheet, onvalue=1, offvalue=0, command=lambda:print_selection(is_all_sheet))
	is_all_sheet_checker.grid(row=0, column=4, sticky="e",pady=2)

	executeButton = Button(execute_frame, text ="Execute", command = execute_change_data)
	executeButton.grid(row=0, column=6, sticky="s",pady=2)


def startup():
	window = create_main_window()
	directory_input = display_directory_input(window)
	change_data_tab, ctr_home_tab = display_tab(window)
	scrollable_frame = display_change_data_content_tab(change_data_tab)
	display_execute_content_tab(window)
	return window, change_data_tab, ctr_home_tab,scrollable_frame,directory_input



