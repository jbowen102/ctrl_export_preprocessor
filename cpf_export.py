import os
import pyautogui as gui


FILE_BUTTON_X_POS=978;
FILE_BUTTON_Y_POS=70+47;

def export_cpfs():
    gui.moveTo(FILE_BUTTON_X_POS, FILE_BUTTON_Y_POS);
    gui.click(); # Press File button in toolbar
    gui.moveRel(29, 20);
    gui.click(); # Press Open
    gui.moveRel(236, 164, .5);
    gui.click(); # Select file to import.
    gui.hotkey("f2"); # Rename shortcut
    gui.hotkey("ctrl", "a"); # Highlight whole filename, including extension
    gui.hotkey("ctrl", "c"); # Copy filename, including extension
    gui.press(["esc"]), # Exit rename
    gui.doubleClick(); # Confirm file to open
    gui.moveTo(FILE_BUTTON_X_POS, FILE_BUTTON_Y_POS, .5);
    gui.click(); # Press File button in toolbar again
    gui.moveRel(65, 113, .2);
    gui.click(); # Click Export
    gui.hotkey("ctrl", "v"); # Paste in imported filename (includes extension)
    gui.press(["backspace"]*4) # Remove extension and 
    gui.press(["Enter"]) # Confirm export filename


if __name__ == "__main__":
    export_cpfs()
    print("done")


