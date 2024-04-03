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
  value = entry_input[0].row

window.mainloop()
