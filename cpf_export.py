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
    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(IMPORT_DIR) # Navigate to import folder.
    gui.press(["enter"])

    # Copy filename to clipboard for use in export.
    gui.click(1207, 209); # Select first file in CPF_DIR to import.
    gui.hotkey("f2"); # "Rename" shortcut
    gui.hotkey("ctrl", "c"); # Copy filename, excluding extension
    gui.press(["esc"]), # Exit rename

    gui.press(["enter"]) # Confirm CPF filename to open.
    time.sleep(1) # Allow time for CPF to open.

    # Export
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    gui.press(["e"]) # Select Export from File menu.
    gui.hotkey("ctrl", "v"); # Paste in imported filename

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(EXPORT_DIR) # Navigate to export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save


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
