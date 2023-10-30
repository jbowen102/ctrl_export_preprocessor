import os
import time

import pyautogui as gui

from dir_names import CPF_DIR, IMPORT_DIR, EXPORT_DIR


# Constants
PROG_POS_X=1433
PROG_POS_Y=547



def export_cpfs():
    gui.click(PROG_POS_X, PROG_POS_Y) # Click on program to bring into focus

    # Open file
    gui.hotkey("ctrl", "o")
    time.sleep(0.2)
    gui.hotkey("ctrl", "l") # Select address bar
    time.sleep(0.2)
    gui.typewrite(IMPORT_DIR) # Navigate to import folder.
    time.sleep(0.2)
    gui.press(["enter"])
    time.sleep(0.2)

    # Copy filename to clipboard for use in export.
    gui.click(1207, 209); # Select first file in CPF_DIR to import.
    gui.hotkey("f2"); # "Rename" shortcut
    time.sleep(0.2)
    gui.hotkey("ctrl", "c"); # Copy filename, excluding extension
    time.sleep(0.2)
    gui.press(["esc"]), # Exit rename
    time.sleep(0.2)

    gui.press(["enter"]) # Confirm CPF filename to open.
    time.sleep(1) # Allow time for CPF to open.

    # Export
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    time.sleep(0.2)
    gui.press(["e"]) # Select Export from File menu.
    time.sleep(0.2)
    gui.hotkey("ctrl", "v"); # Paste in imported filename
    time.sleep(0.2)

    gui.hotkey("ctrl", "l") # Select address bar
    time.sleep(0.1)
    gui.typewrite(EXPORT_DIR) # Navigate to export folder.
    time.sleep(0.1)
    gui.press(["enter"])
    time.sleep(0.2)
    gui.hotkey("alt", "s") # Save

    time.sleep(0.5)
    gui.hotkey("ctrl", "f4") # Close CPF file.


if __name__ == "__main__":
    # Make CPF dirs if any don't exist yet.
    if not os.path.exists(CPF_DIR):
        os.mkdir(CPF_DIR)
    if not os.path.exists(IMPORT_DIR):
        os.mkdir(IMPORT_DIR)
    if not os.path.exists(EXPORT_DIR):
        os.mkdir(EXPORT_DIR)

    input("Ready for GUI interaction?")
    export_cpfs()
    print("GUI interaction done")
