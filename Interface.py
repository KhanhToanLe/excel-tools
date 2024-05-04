import xlsxHelper as xlsx
from tkinter import *
import tkinter.filedialog
from tkinter import ttk
import helper as HELPER
import InputData as data
import traceback
import exception_logger as exlogger

# global variable
input_directory_entry = ""
is_check_all_page = 0
old_choose_value = ""
chose_sheet = 1
notebook_control = None
main_window = ""
process_bar = None
process_bar_frame = None
process_bar_merge = None
process_bar_merge_frame = None

checkbox_frame = None
old_change_data_execute_value = 1
destination_file = None
remove_old_merge_sheet_check_box = None
remove_old_merge_sheet_first = None
language_select_combobox = None


def create_main_window():
    """
    create main window
    """
    main_window = Tk()
    main_window.geometry(HELPER.WINDOW_SIZE)
    main_window.title(HELPER.APP_NAME)
    return main_window


def select_working_directory(input_directory: Entry):
    """
    select working directory button handler
    """
    select_directory = tkinter.filedialog.askdirectory()
    input_directory.delete(0, END)
    input_directory.insert(0, select_directory)


def chose_sheet_data_checker(value):
    global chose_sheet, checkbox_frame
    try:
        select_sheet = int(value)
        checkbox_frame.deselect()
        chose_sheet.config(state="normal")
        chose_sheet.delete(0, END)
        chose_sheet.insert(0, select_sheet)
        return True
    except Exception as ex:
        if value == data.INFINITE_SHEET_CHECKER:
            checkbox_frame.select()
            chose_sheet.config(state="normal")
            chose_sheet.delete(0, END)
            chose_sheet.insert(0, value)
            chose_sheet.config(state="disabled")
            return True
        xlsx.add_message_box("Invalid number sheet", main_window)
        exlogger.write(traceback.format_exc())
        return False


def change_tab_handler(selected_tab):
    global checkbox_frame, is_check_all_page, old_change_data_execute_value, chose_sheet
    if selected_tab == data.CONTROL_HOME_TAB_NAME:
        old_change_data_execute_value = chose_sheet.get()
        checkbox_frame.select()
        is_valid = checker_all_page_handler(IntVar(value=1))
        checkbox_frame.config(state="disabled")
        if is_valid:
            return
    elif selected_tab == data.CHANGE_DATA_TAB_NAME or selected_tab == data.MERGE_SHEET_TAB_NAME and selected_tab ==  data.TRANSLATE_SHEET_TAB_NAME:
        is_valid = chose_sheet_data_checker(old_change_data_execute_value)
        checkbox_frame.config(state="normal")
        if is_valid:
            return


def display_tab(window):
    """
    init and display all tabs
    """
    tab_control = ttk.Notebook(window)
    tab_control.bind("<<NotebookTabChanged>>",
        lambda event: change_tab_handler(tab_control.select()))
    global notebook_control
    notebook_control = tab_control
    tab1 = Frame(tab_control, name="change_data")
    tab2 = Frame(tab_control, name="control_home")
    tab3 = Frame(tab_control, name="merge_sheet")
    tab4 = Frame(tab_control, name='translator')
    tab_control.add(tab1, text='Change Data')
    tab_control.add(tab2, text='Ctrl Home')
    tab_control.add(tab3, text='Merge Sheet')
    tab_control.add(tab4, text="Translate")
    tab_control.pack(expand=1, fill="both")
    return tab1, tab2, tab3,tab4


def display_directory_input(window):
    """
    display directory input field and select working directory button
    """
    pane = Frame(window)
    pane.pack(fill=BOTH)
    Label(pane, text="Directory").pack(side=LEFT, pady=4, padx=4)
    border_color = Frame(pane, background="gray")
    input_directory = Entry(border_color)
    border_color.pack(padx=1, pady=1, fill='x', side=LEFT, expand=TRUE)
    input_directory.pack(padx=1, pady=1, fill='x')
    global input_directory_entry
    input_directory_entry = input_directory
    button = Button(pane, text="Select", background="white", fg="black",
        command=lambda: select_working_directory(input_directory))
    button.pack(side=LEFT, pady=4, padx=12)


def display_change_data_content_tab(tab1):
    """
    display content of content tab
    """
    container = Frame(tab1, background="white")
    height = HELPER.WINDOW_SIZE.split("x")[1]
    canvas = Canvas(container, background="white", height=(int(height)-100))
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
    add_change_data_click_handler(scrollable_frame, 1)
    for i in data.change_data:
        i.frame.pack()
    return scrollable_frame


def reload_change_data_list(scrollable_frame):
    """
    reload change data command in content tab
    """
    for i in scrollable_frame.winfo_children():
        i.pack_forget()
    for i in data.change_data:
        i.frame.pack(fill=X, expand=True)


def remove_change_data(scrollable_frame, added_frame):
    """
    remove a change data command button handler
    """
    data.change_data.remove(added_frame)
    if len(data.change_data) == 0:
        add_change_data_click_handler(scrollable_frame, 1)
    reload_change_data_list(scrollable_frame)


def upper_case_cell(event: Event):
    """
    uppercase value of cell-input's change data command 
    """
    entry = event.widget
    primitive_val = entry.get()
    entry.delete(0, END)
    entry.insert(0, primitive_val.upper())


def add_change_data_click_handler(scrollable_frame, add_item):
    """
    add a change data command button handler
    """
    test = Frame(scrollable_frame)
    cell = Entry(test, font=("Arial", 11, "normal"), width=12)
    cell.bind("<FocusOut>", upper_case_cell)
    cell.pack(side=LEFT)
    value = Entry(test, font=("Arial", 11, "normal"), width=37)
    value.pack(side=LEFT, fill='x', expand=True)
    added_frame = data.InputData(test)
    Button(test, width=4, text="+",
        command=lambda: add_change_data_click_handler(scrollable_frame, 1)).pack(side=LEFT)
    Button(test, width=4, text="-",
        command=lambda: remove_change_data(scrollable_frame, added_frame)).pack(side=LEFT)
    if add_item == 1:
        data.change_data.append(added_frame)
    else:
        insert_after = data.chang_data.index(add_item)
        data.change_data.insert(insert_after, added_frame)
    reload_change_data_list(scrollable_frame)


def checker_all_page_handler(is_check):
    """
    all-page checkbox handler 
    """
    global is_check_all_page, chose_sheet, old_choose_value
    is_check_all_page = is_check.get()
    if (is_check_all_page == 1):
        if chose_sheet.get() != data.INFINITE_SHEET_CHECKER:
            old_choose_value = chose_sheet.get()
        chose_sheet.delete(0, END)
        chose_sheet.insert(0, '∞')
        chose_sheet.config(state="disabled")
    else:
        chose_sheet.config(state="normal")
        chose_sheet.delete(0, END)
        chose_sheet.insert(0, old_choose_value)


def display_change_data_execute_selection(window):
    """
    display execute panel at the button
    """

    execute_frame = Frame(window)
    execute_frame.pack(fill=BOTH, expand=True)

    Label(execute_frame, text="Select Sheet").grid(row=0, column=1)

    select_sheet = Entry(execute_frame, width=4)
    select_sheet.insert(0, 1)
    select_sheet.grid(row=0, column=2, padx=4)
    global chose_sheet, checkbox_frame
    chose_sheet = select_sheet

    my_var = IntVar()
    checkbox_frame = Checkbutton(execute_frame, text='Every Sheet', variable=my_var,
        onvalue=1, offvalue=0, command=lambda: checker_all_page_handler(my_var))
    checkbox_frame.grid(row=0, column=3)
    execute_button = Button(execute_frame, text="Execute",
        command=lambda: execute_button_handler())
    execute_button.grid(row=0, column=4)


def process_bar_increase_val(value):
    """
    change value for creating processBar animation
    """
    if value == 0:
        process_bar['value'] = value
    else:
        if process_bar['value'] < 100:
            process_bar['value'] += value
    main_window.update_idletasks()


def process_bar_merge_increase_val(value):
    """
    change value for creating processBar animation
    """
    global process_bar_merge
    if value == 0:
        process_bar_merge['value'] = value
    else:
        if process_bar_merge['value'] < 100:
            process_bar_merge['value'] += value
    main_window.update_idletasks()


def display_control_home(control_home_tab):
    """
    display control home tab
    """
    global process_bar, process_bar_frame

    process_bar_frame = Frame(control_home_tab, background='white')
    process_bar_frame.pack(fill=BOTH, expand=True)

    Label(process_bar_frame, text="Apply control-home action by pressing execute button\nIt will take a few minutes!",
        height=20, background='white').pack()
    process_bar = ttk.Progressbar(process_bar_frame, orient="horizontal", length=200, mode="determinate",
        takefocus=True, maximum=100)
    process_bar['value'] = 0
    process_bar.pack()


def get_change_data_command():
    """
    combine change data command(<changeData>) to command data(<CommandData>)
    """
    data.command_data = []
    for item in data.change_data:
        current_frame = item.frame
        entries = current_frame.winfo_children()
        data.command_data.append(data.ChangeData(
            entries[0].get(), entries[1].get()))

def display_translate_sheet_tab_content(translate_tab):
    """
    display translate sheet
    """
    translate_frame = Frame(translate_tab,background="white")
    translate_frame.pack(fill=BOTH,expand=True)

    keys_only = [option["key"] for option in HELPER.LANGUAGE_OPTION]
    select_value = StringVar()
    combobox = ttk.Combobox(translate_frame,textvariable= select_value)
    combobox.set("English")
    combobox['values'] = keys_only
    # combobox.set(keys_only[0])
    # combobox.current(0)
    # select_value = keys_only[0]
    combobox.pack(fill=X,expand=True,padx=12)
    global language_select_combobox
    language_select_combobox = combobox
    


def execute_button_handler():
    """
    execute button handler
    """
    global notebook_control, input_directory_entry, chose_sheet, destination_file, remove_old_merge_sheet_check_box, language_select_combobox
    select_tab = notebook_control.select()
    input_directory = input_directory_entry.get()
    chose_sheet_value = chose_sheet.get()
    if select_tab == data.CHANGE_DATA_TAB_NAME:
        get_change_data_command()
        xlsx.change_data(data.command_data, input_directory,
            main_window, chose_sheet_value)
    elif select_tab == data.CONTROL_HOME_TAB_NAME:
        xlsx.control_home(input_directory, chose_sheet_value,
            main_window, process_bar_increase_val)
    elif select_tab == data.MERGE_SHEET_TAB_NAME:
        xlsx.merge_sheet(input_directory, chose_sheet_value, main_window, destination_file.get(
        ), remove_old_merge_sheet_first.get(), process_bar_merge_increase_val)
    elif select_tab == data.TRANSLATE_SHEET_TAB_NAME:
        xlsx.translate_sheet(input_directory,language_select_combobox.get(),chose_sheet_value,main_window,"test")


def choose_destination_file_handler():
    select_file = tkinter.filedialog.askopenfile(
        filetypes=[("Excel files", ".xlsx")])
    global destination_file
    destination_file.delete(0, END)
    value = select_file.name if select_file != None else ""
    destination_file.insert(0, value)


def display_merge_sheet_tab_content(merge_frame):
    merge_container_frame = Frame(merge_frame, background="white")
    merge_container_frame.pack(expand=True, fill=BOTH)
    merge_content_frame = Frame(
        merge_container_frame, background="white", pady=50)
    merge_content_frame.pack()
    Label(merge_content_frame, text="Select directory and destination file before pressing execute button to merge sheet! \nChanging prefix merge-sheet's name in config.json may cause error \nby sheet name duplication", background='white').pack()

    choose_file_frame = Frame(merge_container_frame,
        background="white", padx=12)
    choose_file_frame.pack(fill=X)

    destination_input_frame = Frame(choose_file_frame)
    destination_input_frame.pack(fill=X, expand=True, side=BOTTOM)

    Label(choose_file_frame, text="Destination file:", anchor="e",
        justify=LEFT, background='white').pack(side=LEFT)
    global destination_file
    destination_file = Entry(destination_input_frame, highlightthickness=1)
    destination_file.config(highlightbackground="gray", highlightcolor="gray")
    destination_file.pack(fill=X, expand=True, side=LEFT)
    Button(destination_input_frame, text="Select",
        command=lambda: choose_destination_file_handler()).pack(side=LEFT)

    global remove_old_merge_sheet_check_box, remove_old_merge_sheet_first
    remove_old_merge_sheet_first = IntVar(value=1)
    remove_old_merge_sheet_check_box = Checkbutton(
        choose_file_frame, text="Remove old merge sheet first", variable=remove_old_merge_sheet_first, onvalue=1, offvalue=0, background="white")
    remove_old_merge_sheet_check_box.pack(side=TOP, fill=X, expand=True)

    global process_bar_merge, process_bar_merge_frame

    process_bar_merge_frame = Frame(
        merge_container_frame, background='white', pady=104)
    process_bar_merge_frame.pack(fill=BOTH, expand=True)

    process_bar_merge = ttk.Progressbar(process_bar_merge_frame, orient="horizontal", length=200, mode="determinate",
        takefocus=True, maximum=100)
    process_bar_merge['value'] = 0
    process_bar_merge.pack()


def startup():
    global main_window
    window = create_main_window()
    main_window = window

    display_directory_input(window)
    change_data_tab, ctr_home_tab, merge_tab, translate_tab = display_tab(window)
    scrollable_frame = display_change_data_content_tab(change_data_tab)
    display_control_home(ctr_home_tab)
    display_change_data_execute_selection(window)
    display_merge_sheet_tab_content(merge_tab)
    display_translate_sheet_tab_content(translate_tab)
    return window, change_data_tab, ctr_home_tab, scrollable_frame
