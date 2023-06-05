import openpyxl


# sheet is openpyxl worksheet
def validate_input_file(file):
    # check if column names are in English
    export_wb = openpyxl.load_workbook(file)
    sheet = export_wb.active
    if sheet["A1"].value != "Entry Date":
        raise TypeError
    # may be extended for further validation
    