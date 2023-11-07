import os
import time
import re
import subprocess
import shutil

import colorama
if os.name == "nt":
    # Allows me to test other (non-GUI) features in WSL where pyautogui import fails
    import pyautogui as gui

from dir_names import DIR_REMOTE, \
                      DIR_FIELD_DATA, \
                        DIR_IMPORT_ROOT, DIR_REMOTE_MIRROR, DIR_IMPORT_DATESTAMPED, \
                        DIR_EXPORT


# Constants
PROG_POS_X=1433
PROG_POS_Y=547

DATE_FORMAT = "%Y%m%d"


def find_in_string(regex_pattern, string_to_search, prompt, allow_none=False):
    found = None # Initialize variable for loop
    while not found:
        matches = re.findall(regex_pattern, string_to_search, flags=re.IGNORECASE)
        if len(matches) == 1:
            found = matches[0]
            # loop exits
        elif len(matches) == 0 and allow_none:
            return None
        else:
            print(prompt)
            string_to_search = input("> ")
    return found


def update_import():
    source_dir = DIR_REMOTE_MIRROR
    target_dir = DIR_IMPORT_DATESTAMPED

    for dirpath, dirnames, filenames in os.walk(source_dir):
        for file_name in sorted(filenames):
            filepath = os.path.join(dirpath, file_name)
            item_name = os.path.splitext(file_name)[0]
            ext = os.path.splitext(file_name)[-1]

            if ext.lower() in (".cpf", ".cdf"):
                # Find S/N in filename
                sn_regex = r"(3\d{6}|5\d{6}|8\d{6})"
                # Any "3" or "5" or "8" followed by six more digits
                prompt_str = 'Can\'t parse S/N from import filename "%s". ' \
                                                'Type S/N manually:' % file_name
                serial_num = find_in_string(sn_regex, item_name, prompt_str)

                # Now look for date in remaining string. Will add later if not present.
                remaining_str = item_name.split(serial_num)
                date_found = False
                for substring in remaining_str:
                    # date_regex = r"(20\d{2}(0\d|1[0-2])([0-2]\d|3[0-1]))" # didn't work
                    # Any "20" followed by two digits,
                        # followed by either "0" and a digit or "10", "11", or "12" (months)
                            # followed by either "0", "1", or "2" paired with a digit (days 01-29)
                            # or "30" or "31"
                    date_regex = r"(20\d{2}[0-1]\d[0-3]\d)"
                    # Any "20" followed by two digits,
                        # followed by either "0" or "1" and any digit (months)
                            # followed by either "0", "1", "2", or "3" paired with a digit (days 01-31)

                    prompt_str = 'Found more than one date match in import ' \
                                    'filename "%s". Type manually:' % file_name
                    date_match = find_in_string(date_regex, substring, prompt_str, allow_none=True)

                    if date_match:
                        existing_datestamp = date_match

                        date_found = True
                    else:
                        pass

                if date_found:
                    datestamp = existing_datestamp
                else:
                    # Add datestamp
                    # Find file last-modified time. Precise enough for our needs.
                    mod_date = time.localtime(os.path.getmtime(filepath))

                    mod_date_str = time.strftime(DATE_FORMAT, mod_date)
                    datestamp = mod_date_str

                new_filename = "%s_%s%s" % (datestamp, serial_num, ext)
                new_filepath = os.path.join(target_dir, new_filename)
                if not os.path.exists(new_filepath):
                    shutil.copy2(filepath, new_filepath)


def update_remote_mirror():
    # Sync from remote folder to local one to buffer before processing.
    if os.name == "nt":
        print("Attempting to run robocopy..." + colorama.Fore.YELLOW)
        returncode = subprocess.call(["robocopy", DIR_REMOTE, DIR_REMOTE_MIRROR,
                                                    "/s", "/purge", "/compress"])
                                        # "*.cpf", "/s", "/purge", "/compress"])
        # Removes any extraneous files from local import folder that don't exist in remote.
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders
        print(colorama.Style.RESET_ALL)

        # Check for success
        if returncode < 8:
            # https://superuser.com/questions/280425/getting-robocopy-to-return-a-proper-exit-code
            # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
            print("Sync successful\n")
        else:
            raise Exception("SYNC FAILED")
    elif os.name == "posix":
        print("Attempting to run rsync..." + colorama.Fore.YELLOW)
        # CompProc = subprocess.run(["rsync", "-azivh",
        CompProc = subprocess.run(["rsync", "-azivh", "--delete-before",
            "%s/" % DIR_REMOTE, "%s/" % DIR_REMOTE_MIRROR], stderr=subprocess.STDOUT)
        # Removes any extraneous files from local import folder that don't exist in remote.
        print(colorama.Style.RESET_ALL)

        # Check for success
        if CompProc.returncode == 0:
            print("Sync successful\n")
        else:
            raise Exception("SYNC FAILED")
    else:
        raise Exception("Unrecognized OS type: %s" % os.name)



def convert_file(data_type, source_file_path, target_dir):
    if data_type.lower() == "cpf":
        open_cpf(source_file_path)
        export_cpf(target_dir, filename)


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


def convert_all(file_type, source_dir, dest_dir):
    select_program()
    for filename in sorted(os.listdir(source_dir)):
        filepath = os.path.join(source_dir, filename)
        if (os.path.isfile() and
                    os.path.splitext(filename)[-1].lower() == ".%s" % file_type):
            print("Processing %s..." % filename)
            convert_file(file_type, filepath)
            print("\tdone")
        else:
            # Skip directories
            continue


def create_file_struct():
    # Make field-data dirs if any don't exist yet.
    if not os.path.exists(DIR_FIELD_DATA):
        os.mkdir(DIR_FIELD_DATA)
        print("Created %s" % DIR_FIELD_DATA)
    if not os.path.exists(DIR_IMPORT_ROOT):
        os.mkdir(DIR_IMPORT_ROOT)
        print("Created %s" % DIR_IMPORT_ROOT)
    if not os.path.exists(DIR_IMPORT_DATESTAMPED):
        os.mkdir(DIR_IMPORT_DATESTAMPED)
        print("Created %s" % DIR_IMPORT_DATESTAMPED)
    if not os.path.exists(DIR_EXPORT):
        os.mkdir(DIR_EXPORT)
        print("Created %s" % DIR_EXPORT)


if __name__ == "__main__":

    create_file_struct()

    # Pull from remote CPF dir.
    if os.listdir(DIR_REMOTE_MIRROR):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate local '
                            'import folder from "%s" ? [Y / N]' % DIR_REMOTE)
        run_sync = input("> " + colorama.Style.RESET_ALL)
    else:
        # If DIR_IMPORT_DATESTAMPED empty, don't prompt for sync. Just do it.
        run_sync = "Y"

    if run_sync.upper() == "Y":
        update_remote_mirror()
        update_import()

    else:
        print("Skipping import-dir update from remote.\n")
        # Accept any answer other than Y/y as negative.
        pass

    if os.listdir(DIR_EXPORT):
        # Clear export dir before running?
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT +
                    "Export dir populated. Delete contents before processing? [Y / N]")
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer.upper() == "Y":
            for item in sorted(os.listdir(DIR_EXPORT)):
                os.remove(os.path.join(DIR_EXPORT, item))
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
    convert_all(DIR_IMPORT_DATESTAMPED, DIR_EXPORT, type="cpf")
    print("\nGUI interaction done\n")
