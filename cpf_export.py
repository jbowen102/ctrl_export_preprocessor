import os
import time
import subprocess

import colorama
import pyautogui as gui

from dir_names import CPF_DIR, CPF_DIR_REMOTE, IMPORT_DIR, EXPORT_DIR


# Constants
PROG_POS_X=1433
PROG_POS_Y=547


def update_import_files():
    # Sync from remote folder to local one to buffer before processing.
    if os.name == "nt":
        print("Attempting to run robocopy..." + colorama.Fore.YELLOW)
        returncode = subprocess.call(["robocopy", CPF_DIR_REMOTE, IMPORT_DIR,
                                                        "/purge", "/compress"])
        # Removes any extraneous files from local import folder that don't exist in remote.
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy?redirectedfrom=MSDN
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders
        print(colorama.Style.RESET_ALL)

        # Check for success
        if returncode in [0, 1, 2]:
            # https://superuser.com/questions/280425/getting-robocopy-to-return-a-proper-exit-code
            print("Sync successful\n")
        else:
            raise Exception("SYNC FAILED")
    elif os.name == "posix":
        print("Attempting to run rsync..." + colorama.Fore.YELLOW)
        CompProc = subprocess.run(["rsync", "-azivh", "--delete-before",
                        CPF_DIR_REMOTE, IMPORT_DIR], stderr=subprocess.STDOUT)
        # Removes any extraneous files from local import folder that don't exist in remote.
        print(colorama.Style.RESET_ALL)

        # Check for success
        if CompProc.returncode == 0:
            print("Sync successful\n")
        else:
            raise Exception("SYNC FAILED")
    else:
        raise Exception("Unrecognized OS type: %s" % os.name)


def select_program():
    # Brings 1314 program into focus.
    gui.click(PROG_POS_X, PROG_POS_Y) # Click on program to bring into focus


def open_cpf(file_path):
    # Assumes 1314 program already in focus.
    # Get to import folder
    gui.hotkey("ctrl", "o")

    gui.hotkey("ctrl", "l") # Select address bar

    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    gui.press(["enter"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm CPF filename to open.
    time.sleep(1) # Allow time for CPF to open.


def export_cpf(target_dir, filename):
    xls_filename = os.path.splitext(filename)[0] + ".XLS"

    # Assumes 1314 program already in focus.
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    gui.press(["e"]) # Select Export from File menu.

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(xls_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.2)

    gui.hotkey("ctrl", "f4") # Close CPF file.

    # Check if new file exists in exported location as expected after conversion.
    assert os.path.exists(os.path.join(target_dir, xls_filename)), "Can't confirm output file existence."


def convert_all(import_dir, export_dir):
    import_files = sorted(os.listdir(import_dir))
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
        print("Created %s" % CPF_DIR)
    if not os.path.exists(IMPORT_DIR):
        os.mkdir(IMPORT_DIR)
        print("Created %s" % IMPORT_DIR)

    # Pull from remote CPF dir.
    if os.listdir(IMPORT_DIR):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + "\nUpdate local "
                            "import folder from %s ? [Y / N]" % CPF_DIR_REMOTE)
        run_sync = input("> " + colorama.Style.RESET_ALL)
    else:
        # If IMPORT_DIR empty, don't prompt for sync. Just do it.
       run_sync = "Y"

    if run_sync.upper() == "Y":
        update_import_files()
    else:
        print("Skipping import-dir update from network drive.\n")
        # Accept any answer other than Y/y as negative.
        pass

    if not os.path.exists(EXPORT_DIR):
        os.mkdir(EXPORT_DIR)
        print("Created %s" % EXPORT_DIR)
    elif os.listdir(EXPORT_DIR):
        # Clear export dir before running?
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT +
                    "Export dir populated. Delete contents before processing? [Y / N]")
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer.upper() == "Y":
            for item in sorted(os.listdir(EXPORT_DIR)):
                os.remove(os.path.join(EXPORT_DIR, item))
        else:
            # Accept any answer other than Y/y as negative.
            pass

    # answer = gui.confirm("Ready for GUI interaction?")
    # if answer != "OK":
    #     raise Exception("User canceled.")
    input(colorama.Fore.GREEN + colorama.Style.BRIGHT +
                    "\nReady for GUI interaction?" + colorama.Style.RESET_ALL)
    print() # blank line

    gui.FAILSAFE = True
    # Allows moving mouse to upper-left corner of screen to abort execution.
    gui.PAUSE = 0.2 # 200 ms pause after each command.
    # https://pyautogui.readthedocs.io/en/latest/quickstart.html
    convert_all(IMPORT_DIR, EXPORT_DIR)
    print("\nGUI interaction done\n")
