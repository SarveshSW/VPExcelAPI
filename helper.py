import openpyxl
from openpyxl.utils import get_column_letter
def find_cell_with_value(workbook, sheet_name, target_value):
    # Load the workbook and select the sheet
    sheet = workbook[sheet_name]

    # Iterate through rows and columns to find the target value
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == target_value:
                return (cell.row, cell.column)  # Returns the cell's address (e.g., 'A1')

    return None 