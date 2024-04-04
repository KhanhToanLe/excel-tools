# from openpyxl import *
# import time 
# import re

# path = "sample_file.xlsx"
# wb = load_workbook("sample_file.xlsx")
# sheet = wb.active  


# is_go_to_loop = True
# if sheet.freeze_panes == None:
#     is_go_to_loop = False

# sheet.freeze_panes = 'A2'
# print(sheet.freeze_panes)
# active_cell = "A6"
# sheet.views.sheetView[0].selection[0].activeCell = active_cell
# sheet.views.sheetView[0].selection[0].sqref =  active_cell
# print(sheet.freeze_panes)


# while True and is_go_to_loop:
#     print("loop")
#     matches = re.match(r"([a-zA-Z]+)(\d+)", sheet.freeze_panes)
#     if matches:
#         alphabetic_part = matches.group(1)
#         numeric_part = int(matches.group(2))
#         if numeric_part != 2:
#             numeric_part = numeric_part - 1
#             sheet.freeze_panes = f"{alphabetic_part}{numeric_part}"
#             print(sheet.freeze_panes)
#         else:
#             break

# wb.save("sample_file.xlsx") 


# import xlwings as xw
# excel_app = xw.App(visible=False)

# wb = excel_app.books.open(path)
# ws = wb.sheets[0]
# active_window = wb.app.api.ActiveWindow
# go_to_range = "A1"
# ws.range(go_to_range).select()
# print(active_window.FreezePanes)
# print(active_window.SplitRow)
# print(active_window.SplitColumn)

# wb.save()
# wb.close()

uppercase_alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

row = 27

divide_val = int(row/26)
offset_val = int(row%26)
# if divide_val == 0:
    # do something