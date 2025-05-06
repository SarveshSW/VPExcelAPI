import api

workbook = api.activate_workbook(r"C:\Users\wanzare\Desktop\agl\demo-sarvesh.xlsm")
if workbook is None:
    print("Failed to activate the workbook.")
    exit(1)
print(workbook)
api.load_info("Sheet1")
print("Available Variables in the parametric study: ", api.list_variables())
'''api.set_value("Boiler.Design.SHTempTarg", 534)

api.set_value("Boiler.Design.SHTempTarg", 535)'''
api.set_value("Boiler.Design.SHTempTarg", 535)
api.run_study()
