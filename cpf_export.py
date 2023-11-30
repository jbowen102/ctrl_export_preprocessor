import os
import time
import re
import subprocess
import shutil

from tqdm import tqdm
import colorama
if os.name == "nt":
    # Allows me to test other (non-GUI) features in WSL where pyautogui import fails
    import pyautogui as gui
    gui.FAILSAFE = True
    # Allows moving mouse to upper-left corner of screen to abort execution.
    gui.PAUSE = 0.2 # 200 ms pause after each command.
    # https://pyautogui.readthedocs.io/en/latest/quickstart.html


from dir_names import DIR_REMOTE, \
                      DIR_FIELD_DATA, \
                        DIR_IMPORT_ROOT, DIR_REMOTE_BU, DIR_IMPORT, \
                        DIR_EXPORT


DATE_FORMAT = "%Y%m%d"

CDF_EXPORT_SUFFIX = "_CDF.xlsx"
CPF_EXPORT_SUFFIX = "_cpf.XLS"


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


def datestamp_remote(remote=DIR_REMOTE):
    file_count = sum(len(files) for _, _, files in os.walk(remote))
    # https://stackoverflow.com/questions/35969433/using-tqdm-on-a-for-loop-inside-a-function-to-check-progress
    with tqdm(total=file_count, colour="#05e4ab") as pbar:
        for dirpath, dirnames, filenames in os.walk(remote):
            for file_name in sorted(filenames):
                pbar.update(1)
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
                                        'filename "%s". Type manually (YYYYMMDD ' \
                                                            'format):' % file_name
                        date_match = find_in_string(date_regex, substring, prompt_str, allow_none=True)

                        if date_match:
                            existing_datestamp = date_match
                            date_found = True
                        else:
                            pass

                    if date_found:
                        datestamp = existing_datestamp
                    else:
                        # Find file last-modified time. Precise enough for our needs.
                        mod_date = time.localtime(os.path.getmtime(filepath))

                        # Some files (CDF at least) have bogus mod dates - usually in 1999 or 2000.
                        # In that case, use today's date.
                        if mod_date < time.strptime("20200101", DATE_FORMAT):
                            # Substitute in today's date
                            date_to_use = time.localtime()
                        else:
                            date_to_use = mod_date

                        datestamp = time.strftime(DATE_FORMAT, date_to_use)

                    new_filename = "%s_sn%s%s" % (datestamp, serial_num, ext)
                    new_filepath = os.path.join(dirpath, new_filename)
                    os.rename(filepath, new_filepath)


def sync_from_remote(src, dest, purge=False):

    if os.name=="nt" and purge:
        flag = ["/purge"]
    elif os.name=="posix" and purge:
        flag = ["--delete-before"]
    elif purge:
        raise Exception("Unrecognized OS type: %s" % os.name)
    else:
        flag = []

    if os.name=="nt":
        print("Attempting to run robocopy..." + colorama.Fore.YELLOW)
        returncode = subprocess.call(["robocopy", src, dest, "/s", "/compress"] + flag)
                                        # "*.cpf", "/s", "/purge", "/compress"])
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders

    elif os.name=="posix":
        print("Attempting to run rsync..." + colorama.Fore.YELLOW)
        CompProc = subprocess.run(["rsync", "-azivh"] + flag + ["%s/" % src,
                                        "%s/" % dest], stderr=subprocess.STDOUT)

    print(colorama.Style.RESET_ALL)

    # Check for success
    if (os.name=="nt" and returncode < 8) or (os.name=="posix" and CompProc.returncode == 0):
        # https://superuser.com/questions/280425/getting-robocopy-to-return-a-proper-exit-code
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        print("Sync to '%s' successful\n" % os.path.basename(dest))
    else:
        raise Exception("SYNC to '%s' FAILED" % os.path.basename(dest))


def back_up_remote():
    # Back up remote folder to local one before datestamping files on remote.
    sync_from_remote(DIR_REMOTE, os.path.join(DIR_REMOTE_BU, "mirror"), purge=True)
    # Removes any extraneous files from local import folder that don't exist in remote.

    sync_from_remote(DIR_REMOTE, os.path.join(DIR_REMOTE_BU, "union"))
    # Leaves all in place


def remote_updates():
    # Pull from remote dir.
    if os.listdir(DIR_IMPORT):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nBack up remote "%s" '
                            '(overwrites existing local BU)? [Y / N]' % DIR_REMOTE)
        run_bu = input("> " + colorama.Style.RESET_ALL)
    else:
        # If DIR_IMPORT empty, don't prompt for sync. Just do it.
        run_bu = "Y"

    if run_bu.upper() == "Y":
        print("Backing up ...")
        back_up_remote()
        print("...done")
    else:
        print("Skipping remote BU.")

    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate filenames '
                    'in remote directory (datestamp) "%s"? [Y / N]' % DIR_REMOTE)
    update_remote_filenames = input("> " + colorama.Style.RESET_ALL)
    if update_remote_filenames.upper() == "Y":
        print("Updating remote filenames...")
        datestamp_remote()
        print("...done")
    else:
        # Accept any answer other than Y/y as negative.
        print("Skipping remote-dir filename updates.")

    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate local import '
                                                    'dir from remote? [Y / N]')
    update_import_dir = input("> " + colorama.Style.RESET_ALL)
    if update_import_dir.upper() == "Y":
        print("Updating local files from remote dir...")
        sync_from_remote(os.path.join(DIR_REMOTE, "CDF Files/"), DIR_IMPORT, purge=True)
        sync_from_remote(os.path.join(DIR_REMOTE, "CPF Files/"), DIR_IMPORT)
        print("...done")
    else:
        print("Skipping import-dir update from remote.\n")


def convert_file(source_file_path, target_dir):
    file_type = os.path.splitext(source_file_path)[-1]
    if file_type.lower() == ".cpf":
        open_cpf(source_file_path)
        export_cpf(target_dir, os.path.basename(source_file_path))

    elif file_type.lower() == ".cdf":
        open_cdf(source_file_path)
        export_cdf(target_dir, os.path.basename(source_file_path))


def select_program(filetype):
    # Brings conversion program into focus.
    answer = gui.confirm("Bring %s-conversion GUI into focus, make sure CAPSLOCK is off, then click OK." % filetype.upper())
    if answer == "OK":
        print("\nGUI interaction commencing. Move mouse "
                                    "pointer to upper left of screen to abort.")
    else:
        raise Exception("User canceled.")


def open_cpf(file_path):
    # Assumes 1314 program already in focus.
    # Get to import folder
    gui.hotkey("ctrl", "o")

    gui.hotkey("ctrl", "l") # Select address bar

    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    gui.press(["enter"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm filename to open.
    time.sleep(1) # Allow time for to open.


def export_cpf(target_dir, filename_orig):
    xls_filename = os.path.splitext(filename_orig)[0] + CPF_EXPORT_SUFFIX

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


def open_cdf(file_path):
    # Ensure file is nonzero size. CIT gives error window for empty file.
    if not os.path.getsize(file_path):
        print("\tSkipping %s (empty file)." % os.path.basename(file_path))
        # Skip file
        return

    # Assumes CIT project open and Programmer window open, in focus.
    gui.press(["alt", "f", "i", "c"]) # Import file

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    gui.press(["enter"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm filename to open.
    time.sleep(1) # Allow time for file to open.


def export_cdf(target_dir, filename_orig):
    xlsx_filename = os.path.splitext(filename_orig)[0] + CDF_EXPORT_SUFFIX
    # Assumes CIT project open and Programmer window open, in focus.
    gui.press(["alt", "f", "e", "s"]) # Export spreadsheet

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(xlsx_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.75)

    gui.press(["enter"]) # Click through error

    # Opens .xlsx file at end. Not sure how to suppress.

def convert_all(file_type, source_dir, dest_dir):
    select_program(file_type)
    for filename in tqdm(sorted(os.listdir(source_dir)), colour="cyan"):
        # Check for existing export
        if (os.path.exists(os.path.join(DIR_EXPORT,
                                             os.path.splitext(filename)[0] + CPF_EXPORT_SUFFIX)
                        or os.path.join(DIR_EXPORT,
                                             os.path.splitext(filename)[0] + CDF_EXPORT_SUFFIX)):
            # Skip if already processed this file.
            continue

        filepath = os.path.join(source_dir, filename)
        if (os.path.isfile(filepath) and
                    os.path.splitext(filename)[-1].lower() == ".%s" % file_type):
            print("\nProcessing %s..." % filename)
            convert_file(filepath, dest_dir)
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

    if not os.path.exists(DIR_REMOTE_BU):
        os.mkdir(DIR_REMOTE_BU)
        print("Created %s" % DIR_REMOTE_BU)

    if not os.path.exists(DIR_IMPORT):
        os.mkdir(DIR_IMPORT)
        print("Created %s" % DIR_IMPORT)

    if not os.path.exists(DIR_EXPORT):
        os.mkdir(DIR_EXPORT)
        print("Created %s" % DIR_EXPORT)


if __name__ == "__main__":

    create_file_struct()
    remote_updates()

    if os.listdir(DIR_EXPORT):
        # Clear export dir before running?
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT +
                "Export dir populated. Delete contents before processing? [Y / N]")
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer.upper() == "Y":
            print("Removing files...")
            for item in tqdm(sorted(os.listdir(DIR_EXPORT)), colour="red"):
                os.remove(os.path.join(DIR_EXPORT, item))
            print("...done")
        else:
            # Accept any answer other than Y/y as negative.
            pass

    if os.name == "nt":
        convert_all("cpf", DIR_IMPORT, DIR_EXPORT)
        convert_all("cdf", DIR_IMPORT, DIR_EXPORT)
        print("\nGUI interaction done\n")
    else:
        print("Skipping GUI interaction (requires Windows system.)")
