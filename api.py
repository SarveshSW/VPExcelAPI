import helper
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

variables_coord = None
results_coord = None

sheet_info = {
    "Sheet_Name": None,
    "Variables": [],
    "Results": [],
    "Variables_coord": None,
    "Results_coord": None,
    "Component_ID_coord": None,
    "Base_Unit_coord": None,
    "Base_Value_coord": None,
    "Values_coord": None,
}
def activate_workbook(path_to_workbook):
    """
    Activate the workbook at the specified path.
    
    :param path_to_workbook: Path to the workbook to activate.
    """

    # Load the workbook
    workbook = load_workbook(path_to_workbook)
    
    # Activate the workbook
    workbook.active = workbook.active
    
    return workbook

def load_info(workbook, sheet_name): 
    
    sheet_info['Variables_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Variables")
    sheet_info['Results_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Results")
    sheet_info['Component_ID_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Component ID")
    sheet_info['Base_Unit_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Base Units")
    sheet_info['Base_Value_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Base Value")
    sheet_info['Values_coord'] = helper.find_cell_with_value(workbook, sheet_name, "Value(s)...")
    sheet_info['Sheet_Name'] = sheet_name
    
def list_variables(workbook):
    """
    List all variables in the specified sheet.
    
    :param workbook: The workbook to search in.
    :param sheet_name: The name of the sheet to search in.
    :return: A list of variables found in the sheet.
    """
    
    
    
    # Get the row and column of the "Variables" cell
    start_row = sheet_info["Variables_coord"][0] + 2  # Start from the next row
    start_col = sheet_info["Variables_coord"][1]  # Column of the "Variables" cell
    
    
    # Iterate through rows until we reach an empty cell in the first column
    for row in range(start_row, workbook[sheet_info['Sheet_Name']].max_row + 1):
        cell_value = workbook[sheet_info['Sheet_Name']].cell(row=row, column=start_col).value
        if cell_value is None:
            break  # Stop if we reach an empty cell
        else:
            sheet_info['Variables'].append(cell_value)  # Add variable to the list
    
    return sheet_info['Variables']


def list_results(workbook):
    """
    List all Results in the specified sheet.
    
    :param workbook: The workbook to search in.
    :param sheet_name: The name of the sheet to search in.
    :return: A list of Results found in the sheet.
    """
    
    
    
    # Get the row and column of the "Results" cell
    start_row = sheet_info["Results_coord"][0] + 2  # Start from the next row
    start_col = sheet_info["Results_coord"][1]  # Column of the "Results" cell
    
    
    # Iterate through rows until we reach an empty cell in the first column
    for row in range(start_row, workbook[sheet_info['Sheet_Name']].max_row + 1):
        cell_value = workbook[sheet_info['Sheet_Name']].cell(row=row, column=start_col).value
        if cell_value is None:
            break  # Stop if we reach an empty cell
        else:
            sheet_info['Results'].append(cell_value)  # Add variable to the list
    
    return sheet_info['Results']