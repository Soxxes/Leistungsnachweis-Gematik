import os
import sys
import shutil
import time
import calendar
import string
from copy import copy
from datetime import datetime, date
import locale

import openpyxl
from dotenv import load_dotenv

from utils.interactive import get_input, handle_failed_input
from utils.preprocessing import prepare_df, clean_name
from utils.report import Report
from utils.style import adjust_cell_dimensions, copy_cell_styles, merge_cells
from utils.validation import validate_input_file
from utils.utils import create_folder, load_template


# to get dates and times in German format
locale.setlocale(locale.LC_ALL, "de_DE")

load_dotenv()

CODE_TO_ACTIVITY = {
    "000": os.getenv("000"),
    "001": os.getenv("001"),
    "002": os.getenv("002"),
    "003": os.getenv("003"),
    "004": os.getenv("004"),
    "005": os.getenv("005"),
    "006": os.getenv("006"),
    "007": os.getenv("007"),
    "999": os.getenv("999"),
}
# codes where the user wants store additional information (comments, the text after a ':')
# after the corresponding code
ADDITIONAL_COMMENTS = os.getenv("ADDITIONAL_COMMENTS_FOR_CODES").split(",")

cwd = os.getcwd()
file = get_input(cwd)

# check if the provided Replicon Export has english column names
try:
    validate_input_file(file)
except TypeError:
    handle_failed_input()

# some preprocessing
data = prepare_df(file)

# get template excel sheet for styles
template = load_template(os.path.join(cwd, "Template", "template.xlsx"))
template_sheet = template["template"]
template_cell_dimensions = {
    "widths": {
        col: template_sheet.column_dimensions[col].width for col in string.ascii_uppercase
    },
    "heights": {
        row: template_sheet.row_dimensions[row].height for row in range(1, 60)\
            if template_sheet.row_dimensions[row].height is not None
    }
}
template_merged_cells = template_sheet.merged_cells.ranges

create_folder("output", os.path.join(cwd, "output"))

# get all available clients
client_groups = data.groupby(data["Client Name"])


if __name__ == "__main__":
    start = time.time()
    # client level (folder)
    for i, (client_name, group) in enumerate(client_groups, 1):
        create_folder(str(client_name), os.path.join(cwd, "output", str(client_name)))

        wbs_group = group.groupby(group["Project Code"])
        # wbs code level (folder)
        for wbs, wbs_group in wbs_group:
            project_name = data.loc[data["Project Code"] == wbs, "Project Name"].values[0]
            clean_project_name = clean_name(project_name)
            create_folder(f"{wbs} ({clean_project_name})")

            year_groups = wbs_group.groupby(wbs_group["Entry Date"].dt.year)
            # year level (folder)
            for year, year_group in year_groups:
                create_folder(str(int(year)))

                month_groups = year_group.groupby(year_group["Entry Date"].dt.month)
                # month level (excel file)
                for month, month_group in month_groups:
                    shutil.copy(os.path.join(cwd,
                                             "Template", "template.xlsx"),
                                             f"{calendar.month_name[month]}.xlsx")
                    month_excel = openpyxl.load_workbook(f"{calendar.month_name[month]}.xlsx")

                    employee_groups = month_group.groupby(month_group["Last Name"])
                    # employee level (excel sheet in the file)
                    for employee, employee_group in employee_groups:
                        first_name = employee_group["First Name"].values[0]
                        employee_sheet = month_excel.create_sheet(f"{first_name} {employee}")
                        # call functions to copy style and values
                        merge_cells(employee_sheet, template_merged_cells)
                        copy_cell_styles(template_sheet, employee_sheet)
                        adjust_cell_dimensions(employee_sheet, template_cell_dimensions)
                        report = Report(employee_group,
                                        month,
                                        year,
                                        f"{first_name} {employee}",
                                        clean_project_name)
                        report.fill_worksheet(employee_sheet, CODE_TO_ACTIVITY, ADDITIONAL_COMMENTS)
                        report.fill_header(employee_sheet)

                    month_excel.remove(month_excel["template"])
                    month_excel.save(f"{calendar.month_name[month]}.xlsx")
                    
                os.chdir("..")
            os.chdir("..")
        os.chdir("..")
        print(f"Progress: {round((i/len(client_groups))*100, 2)} %", end="\r")

    delta = round(time.time() - start, 3)

    input(f"Finished Execution in {delta} s. Press any key to exit ...")
