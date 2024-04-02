from openpyxl import load_workbook


def is_cell_frozen(worksheet, cell):
    # Get the SheetView object
    sheet_view = worksheet.sheet_view
    # Check if there are frozen panes
    if sheet_view.pane is not None and sheet_view.pane.state == "frozen":
        # Get the range of frozen panes
        top_left_cell = sheet_view.pane.topLeftCell
        bottom_right_cell = sheet_view.pane.bottomRightCell

        # Convert the cell reference to row and column indices
        top_left_row, top_left_col = cell_to_indices(top_left_cell)
        bottom_right_row, bottom_right_col = cell_to_indices(bottom_right_cell)
        
        # Convert the cell to check to row and column indices
        check_row, check_col = cell_to_indices(cell)

        # Check if the cell falls within the range of frozen panes
        if top_left_row <= check_row <= bottom_right_row and top_left_col <= check_col <= bottom_right_col:
            return True
    return False

# Load the Excel workbook
wb = load_workbook("C:/Users/toanlk/Desktop/ToolsChangeExcel/test.xlsx")
ws = wb.active
# test = 'A1'
print(ws.rows)

# ws.views.sheetView[0].selection[0].activeCell = test
# ws.views.sheetView[0].selection[0].sqref = test
# print(ws.sheet_view.pane.state)
# ws.freeze_panes = test

wb.save("C:/Users/toanlk/Desktop/ToolsChangeExcel/test.xlsx")
# Remember to close the workbook when you're done
wb.close()

