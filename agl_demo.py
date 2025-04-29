import win32com.client as win32 
import time
import matplotlib.pyplot as plt
import numpy as np

xl = win32.Dispatch("Excel.Application")
xl.Visible = True  # Make Excel visible

# Open the workbook
wb1 = xl.Workbooks.Open(r'C:\Users\wanzare\Desktop\agl\demo-sarvesh.xlsm')

sheet = wb1.Sheets("Sheet1") 

variables = []
results = []


fill_var = False
fill_result = False
for row in range(1, sheet.UsedRange.Rows.Count + 1):
    cell = sheet.Cells(row, 1) 
    #print(f"Row {row}, Column A: {cell.Value}")
    if cell.Value == "Variables":
        fill_var = True
    elif cell.Value == "Results":
        fill_var = False
        fill_result = True
    elif cell.Value is None:
        continue

    if fill_var:
        variables.append(cell)
    elif fill_result:
        results.append(cell)

variables = variables[2:]
results = results[2:]



last_val_coord = variables[-1]
last_val_coord_row = last_val_coord.Row
last_val = sheet.Cells(last_val_coord_row,8).Value


start_row = 15
start_col = 9   
if last_val is not None: 
    for i in range(-5,6):
        sheet.Cells(start_row, start_col + i + 5).Value = last_val + i
    wb1.SaveAs(r'C:\Users\wanzare\Desktop\agl\result.xlsm', FileFormat=52)
wb1.Close(SaveChanges=True)

wb = xl.Workbooks.Open(r'C:\Users\wanzare\Desktop\agl\result.xlsm', ReadOnly=False)
addin_name = "VP.ExcelAddIn"  # Corrected COM add-in name
addin_loaded = False

for addin in xl.COMAddIns:
    if addin.progID == addin_name:
        addin.Connect = True  # Enable the add-in
        addin_loaded = True
        print(f"COM Add-in '{addin_name}' loaded successfully.")
        break

if not addin_loaded:
    print(f"Warning: COM Add-in '{addin_name}' not found!")

    
# Inject VBA macro into the first module of the workbook
vba_code = """
Sub Workbook_Open()
    ' Activate the sheet you want to display
    Sheets("VP").Activate
End Sub

Sub CallVSTOMethod()

    Dim addIn As COMAddIn
    Dim parametricStudy As Object
    Set addIn = Application.COMAddIns("VP.ExcelAddin")
    Set parametricStudy = addIn.Object
    parametricStudy.RunStudy
End Sub
"""

module = wb.VBProject.VBComponents.Add(1) 
module.CodeModule.AddFromString(vba_code)
wb.Save()

com_addin = xl.COMAddIns("VP.ExcelAddIn").Object

xl.Application.Run("CallVSTOMethod")

time.sleep(10)

sheet = wb.Sheets("Sheet1")
variables = []
results = []
fill_var = False
fill_result = False
for row in range(1, sheet.UsedRange.Rows.Count + 1):
    cell = sheet.Cells(row, 1) 
    if cell.Value == "Variables":
        fill_var = True
    elif cell.Value == "Results":
        fill_var = False
        fill_result = True
    elif cell.Value is None:
        continue

    if fill_var:
        variables.append(cell)
    elif fill_result:
        results.append(cell)

variables = variables[2:]
results = results[2:]

results_dict = {cell.Value: cell for cell in results}
variables_dict = {cell.Value: cell for cell in variables}

# Access the desired cell using the dictionary
desired_value = "MiscData.HHV Net Heat Rate"
if desired_value in results_dict:
    desired_cell = results_dict[desired_value]
    print(f"Found cell with value '{desired_value}' at address {desired_cell.Address}")
    
    # Find the maximum value in the subsequent columns of the desired row
    row = desired_cell.Row
    max_value = float('-inf')
    max_col = 0
    for col in range(desired_cell.Column + 7, sheet.UsedRange.Columns.Count + 1):
        cell_value = sheet.Cells(row, col).Value
        if cell_value is not None and isinstance(cell_value, (int, float)) and cell_value > max_value:
            max_value = cell_value
            max_col = col
    
    print(f"Maximum value in {row},{max_col}: {max_value}")
    
    # Highlight the entire column in green for all non-None cells
    green_fill = 0x00FF00  # Green color
    for row in range(1, sheet.UsedRange.Rows.Count + 1):
        cell = sheet.Cells(row, max_col)
        if cell.Value is not None:
            cell.Interior.Color = green_fill

else:
    print(f"Value '{desired_value}' not found in the results list.")

# Save the workbook
wb.Save()
temp_data = sheet.Range(f"I{variables_dict['Boiler.Design.SHTempTarg'].Row}:S{variables_dict['Boiler.Design.SHTempTarg'].Row}")

midpoint_temp = temp_data[6].Value
temp_data_diff = [cell.Value - midpoint_temp for cell in temp_data]


heat_rate_data = sheet.Range(f"I{results_dict['MiscData.TCHR'].Row}:S{results_dict['MiscData.TCHR'].Row}")
midpoint_heat_rate = heat_rate_data[6].Value
heat_rate_data_diff = [((cell.Value - midpoint_heat_rate)/midpoint_heat_rate)*100 for cell in heat_rate_data]



# performing a simple linear regression for 1D polynomial
coeffs = np.polyfit(temp_data_diff, heat_rate_data_diff, deg=1)  
poly_eq = np.poly1d(coeffs)  
# Convert lists to NumPy arrays
temp_data_diff = np.array(temp_data_diff)
heat_rate_data_diff = np.array(heat_rate_data_diff)

# Generate regression line data
x_smooth = np.linspace(temp_data_diff.min(), temp_data_diff.max(), 100) 
y_smooth = poly_eq(x_smooth)  # Compute y-values for regression line

# Create a scatter plot
plt.figure(figsize=(6, 4))
plt.scatter(temp_data_diff, heat_rate_data_diff, color='b', label="Heat Rate vs Temp")
plt.plot(x_smooth, y_smooth, linestyle="-", color="r", label="Regression Line")  # Regression line

# Labels and title
plt.xlabel("Temperature Difference (°C)")
plt.ylabel("Heat Rate Difference (%)")
plt.title("Heat Rate Correction Curve Due to Main Steam Temp Change")
plt.legend()
plt.grid(True)

# Save the figure
img_path = r"C:\Users\wanzare\Desktop\agl\heat_rate_plot.png"  # Change path if needed
plt.savefig(img_path, dpi=300)
plt.close()  # Close the figure

sheet.Pictures().Insert(img_path).Select()

# Save and close
wb.Save()
# Close xl
wb.Close(SaveChanges=True)

xl.Quit()