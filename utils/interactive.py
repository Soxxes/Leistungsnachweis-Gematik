import os
import sys
import openpyxl


def get_input(cwd):
    dirs = os.listdir(cwd)
    # check if there are mutiple Replicon Exports in the current directory
    count = 0
    for fname in dirs:
        if "Timesheet Hours" in fname:
            count += 1
            _fname = fname
    # if not, choose the one
    if count == 1:
        file = _fname
    # if multiple, let user choose
    elif count > 1:
        replicon_files = [dirname for dirname in dirs if "Timesheet Hours" in dirname]
        file_map = {i: dirname for i, dirname in enumerate(replicon_files, 1)}
        print("Found the following Replicon Exports:")
        for i, dirname in file_map.items():
            print(f"{i} - {dirname}")

        for _ in range(3):
            choice = input("\nSelect your file: ")
            try:
                file = file_map[int(choice)]
            except ValueError:
                print("[WARNING] Please provide a number.\n")
                continue
            except KeyError:
                print("[WARNING] Please choose one of the numbers displayed. \n")
                continue
            else:
                break
        else:
            input("Received three invalid inputs. Script termintates. Please press any key ...")
            sys.exit()

    # if none, exit
    else:
        input("No Replicon Export found. Press any key to exit ...")
        sys.exit()

    return file

def handle_failed_input():
    print("Ops ... Seems like your Replicon Export has column names in a different language than English.")
    print("You can fix that by exporting your timesheet again or rename the column names.\n")
    column_names = [
        "Entry Date", "First Name",	"Last Name", "Email", "Client Name", "Client Code",
        "Project Name",	"Project Code",	"Task Name", "Task Code", "Hours", "Comments",
        "Supervisor Name (Current)", "Time Off Type", "Work Type (AUT)", "Work Type (DEU)",
        "Work Type (CHE)", "Work Location", "Time Entry Approval Status", "Timesheet Approval Status"
    ]
    column_names_wb = openpyxl.Workbook()
    column_names_ws = column_names_wb.active
    for col, name in enumerate(column_names, 1):
        column_names_ws.cell(row=1, column=col).value = name
    column_names_wb.save("column_names.xlsx")
    print("There should be a file named 'column_names.xlsx' with the correct column names.")
    input("Please change your column names (you can copy and paste them) and re-run the program.\n"\
          + "Press any key to exit ...")
    sys.exit()

def select_client(config: dict):
    clients = list(config.keys())
    client_map = {i: dirname for i, dirname in enumerate(clients, 1)}
    print(f"Found {len(clients)} clients in config file.")
    for i, client in client_map.items():
        print(f"{i} - {client}")

    for _ in range(3):
        choice = input("\nSelect client: ")
        try:
            client = client_map[int(choice)]
        except ValueError:
            print("[WARNING] Please provide a number.\n")
            continue
        except KeyError:
            print("[WARNING] Please choose one of the numbers displayed. \n")
            continue
        else:
            break
    else:
        input("Received three invalid inputs. Script termintates. Please press any key ...")
        sys.exit()

    return client
