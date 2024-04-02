# import need libraries  
from tkinter import *
from openpyxl import *
from InputData import *

# import custom files
import Interface as inf
import helper

# startup app
window, change_data_tab, ctr_home_tab,scrollable_frame = inf.startup()

change_data_value = []

# entry_input = [
#   InputData(
#     Entry(window, width=35),
#     Entry(window, width=35),
#     Entry(window, width=35),
#   ),
#   InputData(
#     Entry(window, width=35),
#     Entry(window, width=35),
#     Entry(window, width=35),
#   ),InputData(
#     Entry(window, width=35),
#     Entry(window, width=35),
#     Entry(window, width=35),
#   )
# ]

# for canva in entry_input:
#   canva.row.pack();
#   canva.column.pack();
#   canva.value.pack();
  

# a=Entry(window, width=35)
# a.pack()

# b=Entry(window, width=35)
# b.pack()

# c=Entry(window, width=35)
# c.pack()

def set_value_test():
  # wb = load_workbook(filename = "first_test.xlsx")
  # active_sheets = wb.active
  # input_cell = f"{a.get()}{b.get()}";
  # print(input_cell);
  # active_sheets[input_cell] = c.get()
  # wb.save("first_test.xlsx")
  value = entry_input[0].row
  print(value.get())


# AddButton = Button(window, text ="Add value", command = lambda:inf.add_change_data_click_handler(scrollable_frame))
# AddButton.place(x=4,y=60)

# AddValue = Button(window, text ="Add Row", command = inf.add_row_value_input_handler)
# AddValue.place(x=60,y=60)
executeButton = Button(window, text ="Add value", command = lambda:inf.add_change_data_click_handler(change_data_tab))
executeButton.pack(side=BOTTOM)

window.mainloop()