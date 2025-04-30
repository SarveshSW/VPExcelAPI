import api

workbook = api.activate_workbook(r"C:\Users\wanzare\Desktop\agl\demo-sarvesh.xlsm")
api.load_info(workbook, "Sheet1")
print("Available Variables in the parametric study: ", api.list_variables(workbook))
print("Available Variables in the parametric study: ", api.list_results(workbook))
