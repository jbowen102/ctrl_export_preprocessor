import os
import time
import re
import subprocess
import shutil

from tqdm import tqdm
import colorama
from xlsxwriter.workbook import Workbook
if os.name == "nt":
    # Allows me to test other (non-GUI) features in WSL where pyautogui import fails
    import pyautogui as gui
    gui.FAILSAFE = True
    # Allows moving mouse to upper-left corner of screen to abort execution.
    gui.PAUSE = 0.5 # 500 ms pause after each command.
    # https://pyautogui.readthedocs.io/en/latest/quickstart.html


import fix_cpf_export_format as fixcpf
from dir_names import DIR_REMOTE_SRC, \
                      DIR_FIELD_DATA, \
                        DIR_IMPORT_ROOT, DIR_REMOTE_BU, DIR_IMPORT, \
                        DIR_EXPORT, DIR_EXPORT_BUFFER, \
                      DIR_REMOTE_SHARE, \
                      ERROR_HISTORY_SAVE_IMG


DATE_FORMAT = "%Y%m%d"

CDF_EXPORT_SUFFIX = "_CDF.xlsx"

CPF_PARAM_EXPORT_SUFFIX = "_cpf-params.tsv"
CPF_FAULT_EXPORT_SUFFIX = "_cpf-faults.tsv"
CPF_COMBINED_EXPORT_SUFFIX = "_cpf.xlsx"

ERROR_HISTORY_SAVE_BUTTON_LOC = None # Will be modified below


def find_in_string(regex_pattern, string_to_search, prompt, date_target=False, allow_none=False):
    found = None # Initialize variable for loop
    while not found:
        matches = re.findall(regex_pattern, string_to_search, flags=re.IGNORECASE)
        if len(matches) == 1 and date_target:
            # If looking for a date, check for valid date value (regex doesn't fully validate)
            # print("\t\tmatches[0]: " + matches[0]) # DEBUG
            try:
                time.strptime(matches[0], DATE_FORMAT)
            except ValueError:
                # Fall through to prompt user for manual entry.
                pass
            else:
                # Valid date
                found = matches[0]
                # print("\t\t%s: found %s (date that passed both checks)" % (string_to_search, found)) # DEBUG
                return found
        elif len(matches) == 1:
            found = matches[0]
            return found
            # loop exits
        elif len(matches) == 0 and allow_none:
            # print("\t\t%s: no matches; returning None" % string_to_search) # DEBUG
            return None

        # No matches, or invalid date found:
        print(prompt)
        string_to_search = input("> ")


def datestamp_remote(remote=DIR_REMOTE_SRC):
    while not os.path.exists(remote):
        # Prompt user to mount network drives if not found.
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\n"%s" not found. Mount '
                                'and press Enter to try again.' % src)
        input("> " + colorama.Style.RESET_ALL)

    # Keep track of renames for display later.
    old_names = []
    new_names = []

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
                    # print("\n\tS/N:") # DEBUG
                    serial_num = find_in_string(sn_regex, item_name, prompt_str)
                    # Now look for date in remaining string. Will add later if not present.
                    # print("\tReceived %s as S/N back from find_in_string()" % serial_num) # DEBUG
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
                                # followed by either "0", "1", "2", or "3" paired with any digit (days 01-31)
                        # Could catch some invalid dates like 20231131. Further validated below in find_in_string()

                        prompt_str = 'Can\'t find single valid date match in import ' \
                                        'filename "%s". Type manually (YYYYMMDD ' \
                                                            'format):' % file_name
                        # print("\tDate:") # DEBUG
                        date_match = find_in_string(date_regex, substring,
                                    prompt_str, date_target=True, allow_none=True)
                        # print("\tReceived %s as date back from find_in_string()" % date_match) # DEBUG

                        if date_match is not None:
                            date_found = True
                            existing_datestamp = date_match
                            break

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

                    if file_name != new_filename:
                        old_names.append(file_name)
                        new_names.append(new_filename)
                # input("> ") # DEBUG

    print("Renames:")
    if len(old_names) > 0:
        for i, name in enumerate(old_names):
            print(colorama.Fore.MAGENTA + "\t%s\t->\t%s" % (old_names[i], new_names[i]))
    else:
        print(colorama.Fore.MAGENTA + "[None]")
    input(colorama.Fore.GREEN + colorama.Style.BRIGHT + "\nPress Enter to continue")
    print(colorama.Style.RESET_ALL)


def sync_remote(src, dest, purge=False, multilevel=True):
    if not os.path.exists(src):
        raise Exception("Can't find src dir '%s'" % src)
    if not os.path.exists(dest):
        raise Exception("Can't find dest dir '%s'" % dest)

    flags = []
    if multilevel and os.name=="nt":
        flags.append("/s")
    elif not multilevel and os.name=="posix":
        flags.extend(["-f", "- /*/"])
        # https://superuser.com/questions/436070/rsync-copying-directory-contents-non-recursively

    if purge and os.name=="nt":
        flags.append("/purge")
    elif purge and os.name=="posix":
        flags.append("--delete-before")

    if os.name=="nt":
        print("Attempting to run robocopy..." + colorama.Fore.YELLOW)
        returncode = subprocess.call(["robocopy", src, dest, "/compress"] + flags)
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders

    elif os.name=="posix":
        print("Attempting to run rsync..." + colorama.Fore.YELLOW)
        CompProc = subprocess.run(["rsync", "-azivh"] + flags + ["%s/" % src,
                                        "%s/" % dest], stderr=subprocess.STDOUT)

    print(colorama.Style.RESET_ALL)

    # Check for success
    if (os.name=="nt" and returncode < 8) or (os.name=="posix" and CompProc.returncode == 0):
        # https://superuser.com/questions/280425/getting-robocopy-to-return-a-proper-exit-code
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        print("Sync to '%s' successful\n" % os.path.basename(dest))
    else:
        raise Exception("SYNC to '%s' FAILED" % os.path.basename(dest))


def back_up_remote(src=DIR_REMOTE_SRC, dest_root=DIR_REMOTE_BU):
    if not os.path.exists(src):
        raise Exception("Can't find src dir '%s'" % src)
    if not os.path.exists(dest_root):
        raise Exception("Can't find dest_root dir '%s'" % dest_root)

    # Back up remote source contents before datestamping files on remote.
    sync_remote(src, os.path.join(dest_root, "mirror"), purge=True)
    # Removes any extraneous files from local import folder that don't exist in remote.

    sync_remote(src, os.path.join(dest_root, "union"))
    # Leaves all in place


def remote_updates(src=DIR_REMOTE_SRC, dest=DIR_IMPORT):
    if not os.path.exists(dest):
        raise Exception("Can't find dest dir '%s'" % dest)

    # Pull from remote dir.
    if os.listdir(dest):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nBack up remote "%s" '
                            '(overwrites existing local BU)? [Y / N]' % src)
        run_bu = input("> " + colorama.Style.RESET_ALL)
    else:
        # If dest empty, don't prompt for sync. Just do it.
        run_bu = "Y"

    if run_bu.upper() == "Y":
        print("Backing up ...")
        back_up_remote()
        print("...done")
    else:
        print("Skipping remote BU.")

    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate filenames '
                    'in remote directory (datestamp) "%s"? [Y / N]' % src)
    update_remote_filenames = input("> " + colorama.Style.RESET_ALL)
    if update_remote_filenames.upper() == "Y":
        print("Updating remote filenames...")
        datestamp_remote()
        print("...done")

        # Also back up to shared folder for reference.
        print("Syncing source files to shared folder...")
        sync_remote(DIR_REMOTE_SRC, os.path.join(DIR_REMOTE_SHARE, "Raw"), purge=True)
        print("...done")
    else:
        # Accept any answer other than Y/y as negative.
        print("Skipping remote-dir filename updates.")

    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate local import '
                                                    'dir from remote? [Y / N]')
    update_import_dir = input("> " + colorama.Style.RESET_ALL)
    if update_import_dir.upper() == "Y":
        print("Updating local files from remote dir...")
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CDF Files/"), DIR_IMPORT, purge=True)
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CPF Files/"), DIR_IMPORT)
        print("...done")
    else:
        print("Skipping import-dir update from remote.\n")


def convert_file(cxf_path, target_dir, temp_dir=DIR_EXPORT_BUFFER):
    if not os.path.exists(cxf_path):
        raise Exception("Can't find src file '%s'" % cxf_path)
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    file_type = os.path.splitext(cxf_path)[-1]
    cxf_name = os.path.basename(cxf_path)

    if file_type.lower() == ".cpf":
        cpf_open = False
        # Open CPF in GUI and export parameters if export doesn't exist already.
        cpf_param_export_filename = os.path.splitext(cxf_name)[0] + CPF_PARAM_EXPORT_SUFFIX
        if not os.path.exists(os.path.join(temp_dir, cpf_param_export_filename)):
            cpf_open = open_cpf(cxf_path)
            cpf_params_path = export_cpf_params(temp_dir, cpf_param_export_filename)
        else:
            cpf_params_path = os.path.join(temp_dir, cpf_param_export_filename)

        # Open CPF in GUI and export faults if export doesn't exist already.
        cpf_fault_export_filename = os.path.splitext(cxf_name)[0] + CPF_FAULT_EXPORT_SUFFIX
        if not os.path.exists(os.path.join(temp_dir, cpf_fault_export_filename)):
            if not cpf_open:
                cpf_open = open_cpf(cxf_path)

            # Export faults
            cpf_faults_path = export_cpf_faults(temp_dir, cpf_fault_export_filename)
        else:
            cpf_faults_path = os.path.join(temp_dir, cpf_fault_export_filename)

        # Combine both tsvs to single export file.
        cpf_combined_export_filename = os.path.splitext(cxf_name)[0] + CPF_COMBINED_EXPORT_SUFFIX
        cpf_combined_export_path = os.path.join(target_dir, cpf_combined_export_filename)
        fixcpf.combine_param_and_fault_export(cpf_params_path, cpf_faults_path, cpf_combined_export_path)

    elif file_type.lower() == ".cdf":
        open_cdf(cxf_path)

        cdf_export_filename = os.path.splitext(cxf_name)[0] + CDF_EXPORT_SUFFIX
        export_path = export_cdf(target_dir, cdf_export_filename)


def select_program(filetype):
    # Brings conversion program into focus.
    answer = gui.confirm("Bring %s-conversion GUI into focus, make sure CAPSLOCK is off, then click OK." % filetype.upper())
    if answer == "OK":
        print(colorama.Fore.RED + colorama.Style.BRIGHT + "\nGUI interaction "
                        "commencing (%s). Move mouse pointer to upper left of "
                                        "screen to abort." % filetype.upper())
        print(colorama.Style.RESET_ALL)
    else:
        raise Exception("User canceled.")


def open_cpf(file_path):
    if not os.path.exists(file_path):
        raise Exception("Can't find file_path '%s'" % file_path)

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

    return True


def export_cpf_params(target_dir, output_filename):
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes 1314 program already in focus.
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    gui.press(["e"]) # Select Export from File menu.

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.2)

    # Check if new file exists in exported location as expected after conversion.
    export_path = os.path.join(target_dir, output_filename)
    assert os.path.exists(export_path), "Can't confirm output file existence."
    return export_path


def export_cpf_faults(target_dir, output_filename):
    global ERROR_HISTORY_SAVE_BUTTON_LOC # Allow modification of global variable

    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes 1314 program already in focus.

    gui.hotkey("ctrl", "4") # Diagnostics tab

    # Click on Save button inside Error History tab (different than Ctrl+S save)
    # Use previously-found button if coordinates stored already.
    if ERROR_HISTORY_SAVE_BUTTON_LOC is None:
        loc_tuple = gui.locateCenterOnScreen(ERROR_HISTORY_SAVE_IMG)
        if loc_tuple is None:
            pass
            # raise Exception("Can't find Error History save button.")
        ERROR_HISTORY_SAVE_BUTTON_LOC = loc_tuple # Update global variable.
    else:
        loc_tuple = ERROR_HISTORY_SAVE_BUTTON_LOC
    # loc_tuple = gui.locateCenterOnScreen(ERROR_HISTORY_SAVE_IMG)
    # if loc_tuple is None:
    #     # If no faults present in CPF, Save button will be absent.
    #     print("\nNo faults present in %s" % output_filename)
    #     return None
    # # Need to improve robustness of button ID. Above false-trips when Save button actually present

    x, y = loc_tuple
    gui.click(x, y)

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.2)
    gui.hotkey("ctrl", "f4") # Close CPF file.

    # Check if new file exists in exported location as expected after conversion.
    export_path = os.path.join(target_dir, output_filename)
    assert os.path.exists(export_path), "Can't confirm output file existence."
    return export_path


def open_cdf(file_path):
    if not os.path.exists(file_path):
        raise Exception("Can't find file_path '%s'" % file_path)

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


def export_cdf(target_dir, output_filename):
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes CIT project open and Programmer window open, in focus.
    gui.press(["alt", "f", "e", "s"]) # Export spreadsheet

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.75)

    gui.press(["enter"]) # Click through error

    # Opens .xlsx file at end. Not sure how to suppress.

    return os.path.join(target_dir, output_filename)


def convert_all(file_type, source_dir, dest_dir):
    if not os.path.exists(source_dir):
        raise Exception("Can't find source_dir '%s'" % source_dir)
    if not os.path.exists(dest_dir):
        raise Exception("Can't find dest_dir '%s'" % dest_dir)

    select_program(file_type)
    file_list = [x for x in sorted(os.listdir(source_dir)) if x.lower().endswith(file_type)]
    for filename in tqdm(file_list, colour="cyan"):
        # Check for existing export
        if file_type == "cpf" and (os.path.exists(os.path.join(DIR_EXPORT,
                            os.path.splitext(filename)[0] + CPF_COMBINED_EXPORT_SUFFIX))):
            # Skip if already processed this file.
            tqdm.write("Already processed %s" % os.path.basename(filename)) # DEBUG
            continue
        elif file_type == "cdf" and (os.path.exists(os.path.join(DIR_EXPORT,
                            os.path.splitext(filename)[0] + CDF_EXPORT_SUFFIX))):
            # Skip if already processed this file.
            tqdm.write("Already processed %s" % os.path.basename(filename)) # DEBUG
            continue

        filepath = os.path.join(source_dir, filename)
        if (os.path.isfile(filepath) and
                    os.path.splitext(filename)[-1].lower() == ".%s" % file_type):
            try:
                convert_file(filepath, dest_dir)
            except Exception as exception_text:
                print(colorama.Fore.CYAN + colorama.Style.BRIGHT)
                print("\nEncountered exception processing %s" % filename + colorama.Style.RESET_ALL)
                print(exception_text)
                print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
                print("Proceed to convert tsv-format CPF exports that were processed already? [Y/N]")
                answer = input("> " + colorama.Style.RESET_ALL)
                if answer.upper() == "Y":
                    break
                else:
                    raise Exception(exception_text)
            else:
                tqdm.write("Processed %s" % filename)

        else:
            # Skip directories
            continue


def convert_cpfs_in_export(dir_path):
    """Convert CPF exports (.XLS extension but TSV format) to true Excel format.
    Run as batch to catch any exports that didn't get converted and to delete
    old pre-converted exports lingering in export folder."""
    if not os.path.exists(dir_path):
        raise Exception("Can't find dir_path '%s'" % dir_path)

    print("\nConverting CPF exports from tsv format (named .XLS) to .xslx (in dir "
                                                    "\n\t\"%s\")..." % dir_path)
    try:
        fixcpf.convert_all_param_exports(dir_path, check_xls=False)
        print("...done")
    except PermissionError:
        # Gets a PermissionError if running on PowerShell most of the time.
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
        input("\nEncountered permission error in removing CPF tsv files.\n"
                        "Press Enter to continue to next part of program.")
        print(colorama.Style.RESET_ALL)


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
        try:
            convert_all("cpf", DIR_IMPORT, DIR_EXPORT)
            convert_all("cdf", DIR_IMPORT, DIR_EXPORT)
            print(colorama.Fore.RED + colorama.Style.BRIGHT + "\nGUI interaction done\n")
            print(colorama.Style.RESET_ALL)
        except gui.FailSafeException:
            print(colorama.Fore.RED + colorama.Style.BRIGHT + "\n\nUser canceled GUI interaction.")
            print(colorama.Style.RESET_ALL)
            time.sleep(3)
            # If user terminates GUI interraction, continue running below.
            pass
    else:
        print(colorama.Fore.RED + colorama.Style.BRIGHT + "Skipping GUI interaction (requires Windows system.)")
        print(colorama.Style.RESET_ALL)

    print("Syncing processed files to shared folder...")
    sync_remote(DIR_EXPORT, os.path.join(DIR_REMOTE_SHARE, "Converted"), purge=True, multilevel=False)
    print("...done")