import xlwings as xw
from openpyxl import *
from xlsxHelper import *
wb = xw.Book('test/freeze_both.xlsx')

sheet = wb.sheets[0]
cur_sheet = sheet["A1"]
test = parent_of_merged_cell(cur_sheet,sheet)
print(test)

# #copy within the same sheet
# sheet.api.Copy(Before=sheet.api)

# #copy to a new workbook
# sheet.api.Copy()

# app = xw.App(visible=True)
# book = xw.Book("test/freeze_both.xlsx")
# path = "test/freeze_both.xlsx"
# book2 = xw.Book("test/freeze_column.xlsx")
# path1 = "test/freeze_column.xlsx"
# sheet2 = book2.sheets[0]
# sheet = book.sheets[0]

# print(type(sheet.api))
# sheet.api.Copy(After=sheet2.api)

# book.save()
# book2.save()
# app.quit()
# try:
#     app = xw.App(visible=True)

#     first_book = xw.Book(path)
#     second_book = xw.Book(path1)
#     first_book.sheets[0]['A1'].value = 'some value'
    

#     # Copy to same Book with the default location and name
#     first_book.sheets[0].copy()
#     print(first_book.sheets[0].name)
#     # Copy to same Book with custom sheet name
#     first_book.sheets[0].copy(name='copied2')

#     # Copy to second Book requires to use before or after
#     first_book.sheets[0].copy(after=second_book.sheets[0])
# except Exception as exception:
#     print(exception)
# finally:
#     app.quit()



# f = open("test.xlsx", "w")
# f.close()
# directory = ""
# files = get_files(directory)
# wb = load_workbook(filename = file)