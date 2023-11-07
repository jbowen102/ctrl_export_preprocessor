import os
import time
import re
import subprocess

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


def datestamp_filenames(target_directory):
    items = sorted(os.listdir(target_directory))
    for n, file_name in enumerate(items):
        filepath = os.path.join(target_directory, file_name)
        item_name = os.path.splitext(file_name)[0]
        ext = os.path.splitext(file_name)[-1]

        # Check for date already present in filename
        sn_regex = r"(3\d{6}|5\d{6}|8\d{6})"
        # Any "3" or "5" or "8" followed by six more digits
        sn_matches = re.findall(sn_regex, item_name, flags=re.IGNORECASE)
        assert not len(sn_matches) > 1, 'More than one S/N match found in import filename "%s". Unhandled exception.' % file_name
        assert len(sn_matches) == 1, 'No S/N match found in import filename "%s". Unhandled exception.' % file_name
        serial_num = sn_matches[0]

        # Now look for date in remaining string. Will add later if not present.
        # Does not validate any existing datestamp in filename.
        remaining_str = item_name.split(serial_num)
        date_found = False
        for substring in remaining_str:
            if len(substring) >= len("20230101"): # long enough to be a date.
                # date_regex = r"(20\d{2}(0\d|1[0-2])([0-2]\d|3[0-1]))" # didn't work
                # Any "20" followed by two digits,
                    # followed by either "0" and a digit or "10", "11", or "12" (months)
                        # followed by either "0", "1", or "2" paired with a digit (days 01-29)
                        # or "30" or "31"

                date_regex = r"(20\d{2}[0-1]\d[0-3]\d)"
                # Any "20" followed by two digits,
                    # followed by either "0" or "1" and any digit (months)
                        # followed by either "0", "1", "2", or "3" paired with a digit (days 01-31)
                date_matches = re.findall(date_regex, substring, flags=re.IGNORECASE)

                if len(date_matches) == 1:
                    existing_datestamp = date_matches[0]
                    date_found = True
                elif len(date_matches) > 1:
                    raise Exception("More than one date match found in import "
                                "filename %s. Unhandled exception" % item_name)
                else:
                    pass
            else:
                pass

        if date_found:
            datestamp = existing_datestamp
        else:
            # Add datestamp
            # Find file last-modified time. Precise enough for our needs.
            mod_date = time.strftime(DATE_FORMAT, time.localtime(os.path.getmtime(filepath)))
            datestamp = mod_date

        new_filename = "%s_%s%s" % (datestamp, serial_num, ext)
        assert os.path.exists(filepath), "File not found for rename: %s" % filepath
        os.rename(filepath, os.path.join(target_directory, new_filename))


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


def convert_all(DIR_IMPORT_DATESTAMPED, DIR_EXPORT):
    import_files = sorted(os.listdir(DIR_IMPORT_DATESTAMPED))
    for n, filename in enumerate(import_files):
        select_program()
        if (os.path.isfile(os.path.join(DIR_IMPORT_DATESTAMPED, filename)) and
                            os.path.splitext(filename)[-1].lower() == ".cpf"):
            print("Processing %s..." % filename)
            open_cpf(os.path.join(DIR_IMPORT_DATESTAMPED, filename))
            export_cpf(DIR_EXPORT, filename)
            print("\tdone")
        else:
            # Skip directories and non-CPFs
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
        # datestamp_filenames()
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
    convert_all(DIR_IMPORT_DATESTAMPED, DIR_EXPORT)
    print("\nGUI interaction done\n")
