import os
import sys
import yaml
import logging

import openpyxl


def create_folder(name, path=None):
    try:
        os.mkdir(name)
        logging.info(f"Created new folder {name}.")
    except FileExistsError:
        logging.info(f"Skipped folder creation of folder '{name}' since it already exists.")
        pass
    # change back to working directory
    if path is not None:
        os.chdir(path)
        logging.info(f"Changed directory to: '{path}'.")
    else:
        os.chdir(name)
        logging.info(f"Changed directory to: '{name}'.")

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
        with open(path, "r", encoding="utf-8") as ymlfile:
            config = yaml.load(ymlfile, Loader=yaml.SafeLoader)
        return config
    except FileNotFoundError:
        print("[ERROR] Missing config file in template folder.")
        input("Aborting. Press any key to exit ...")
        sys.exit()
