import os
import time

import pyautogui as gui

from dir_names import CPF_DIR, IMPORT_DIR, EXPORT_DIR


# Constants
PROG_POS_X=1433
PROG_POS_Y=547



def select_program():
    # Brings 1314 program into focus.
    gui.click(PROG_POS_X, PROG_POS_Y) # Click on program to bring into focus


def open_cpf(file_path):
    # Assumes 1314 program already in focus.
    # Get to import folder
    gui.hotkey("ctrl", "o")
    time.sleep(0.2)

    gui.hotkey("ctrl", "l") # Select address bar
    time.sleep(0.2)

    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    time.sleep(0.2)
    gui.press(["enter"])
    time.sleep(0.2)

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm CPF filename to open.
    time.sleep(1) # Allow time for CPF to open.


def export_cpf(target_dir, filename):
    xls_filename = os.path.splitext(filename)[0] + ".XLS"

    # Assumes 1314 program already in focus.
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    time.sleep(0.2)
    gui.press(["e"]) # Select Export from File menu.
    time.sleep(0.2)

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(xls_filename)
    time.sleep(0.2)

    gui.hotkey("ctrl", "l") # Select address bar
    time.sleep(0.1)
    gui.typewrite(target_dir) # Navigate to target export folder.
    time.sleep(0.1)
    gui.press(["enter"])
    time.sleep(0.2)
    gui.hotkey("alt", "s") # Save

    time.sleep(0.5)
    gui.hotkey("ctrl", "f4") # Close CPF file.

    # Check if new file exists in exported location as expected after conversion.
    assert os.path.exists(os.path.join(target_dir, xls_filename)), "Can't \
                                                confirm output file existence."


def convert_all(import_dir, export_dir):
    import_files = os.listdir(import_dir)
    for n, filename in enumerate(import_files):
        select_program()
        if (os.path.isfile(os.path.join(import_dir, filename)) and
                            os.path.splitext(filename)[-1].lower() == ".cpf"):
            print("Processing %s..." % filename)
            open_cpf(os.path.join(import_dir, filename))
            export_cpf(export_dir, filename)
            print("\tdone")
        else:
            # Skip directories and non-CPFs
            continue


if __name__ == "__main__":
    # Make CPF dirs if any don't exist yet.
    if not os.path.exists(CPF_DIR):
        os.mkdir(CPF_DIR)
    if not os.path.exists(IMPORT_DIR):
        os.mkdir(IMPORT_DIR)
    if not os.path.exists(EXPORT_DIR):
        os.mkdir(EXPORT_DIR)

    input("Ready for GUI interaction?")
    print()
    convert_all(IMPORT_DIR, EXPORT_DIR)
    print("GUI interaction done\n")
