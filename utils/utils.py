import os
import sys
import openpyxl


def create_folder(name, path_to_go=None):
    try:
        os.mkdir(name)
    except FileExistsError:
        pass
    # change back to working directory
    if path_to_go is not None:
        os.chdir(path_to_go)
    else:
        os.chdir(name)

# path should be os pathlike object
def load_template(path):
    try:
        template = openpyxl.load_workbook(path)
        return template
    except FileNotFoundError:
        print("[ERROR] No template file found. Please provide a template file for your 'Leistungsnachweis'.")
        print("Please create a folder 'Template' with a 'template.xlsx' file inside.")
        input("Aborting. Press any key to exit ...")
        sys.exit()
