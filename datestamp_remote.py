import os.path


import cpf_export as cpf
from dir_names import DIR_REMOTE_SRC, DIR_REMOTE_SHARE


def wait_for_input():
    input("\nEnd of Script. Press Enter to finish and close.")

# Standalone script to be run automatically each day.
try:
    print("Updating remote filenames...")
    cpf.datestamp_remote()
    print("...done")

    # Also back up to shared folder for reference. 
    print("Syncing source files to shared folder...")
    cpf.sync_remote(DIR_REMOTE_SRC, os.path.join(DIR_REMOTE_SHARE, "Raw"), purge=True)
    print("...done")

    wait_for_input()

except Exception as exception_text:
    print(exception_text)
    print("\n" + "*"*10 + "\nException encountered\n" + "*"*10)
    wait_for_input()
