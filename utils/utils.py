import os
import sys
import openpyxl
import yaml


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
def load_template(path, client, config):
    try:
        template = openpyxl.load_workbook(path)
        skip_style = config.get("Clients").get(client).get("skip_style")
        if not skip_style:
            template_sheet = template[config.get("Clients").get(client).get("template_sheet_name")]
        else:
            template_sheet = None
        return template, template_sheet, skip_style
    except FileNotFoundError:
        print("[ERROR] No template file found. Please provide a template file for your 'Leistungsnachweis'.")
        print("Please create a folder 'Template' with a 'template.xlsx' file inside.")
        input("Aborting. Press any key to exit ...")
        sys.exit()

def load_config(path):
    try:
        with open(path, "r") as ymlfile:
            config = yaml.load(ymlfile, Loader=yaml.SafeLoader)
        return config
    except FileNotFoundError:
        print("[ERROR] Missing config file in template folder.")
        input("Aborting. Press any key to exit ...")
        sys.exit()
