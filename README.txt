Copy this .exe to a folder that contains .XLS exports from CPFs.
When double-clicked, the program will translate all the .XLS exports from TSV (tab-separated value)
format to .xlsx. The export from a CPF is actually formatted as a TSV (not XLS),
but it's named w/ .XLS to get Excel to open it.
The benefit is you can now reference the data in those exports from another spreadsheet (like with the
CPF fault aggregator spreadsheet).
Otherwise you'd have to open and save each one to actually be formatted as .XLS or .xlsx for the other
spreadsheet to be able to reference cells within it.

To get the CPF fault aggregator spreadsheet to read values from the other spreadsheets, it seems
you have to have the source files open. Once all the data appears in the sheet though, you can copy
the whole thing and paste it over itself ("paste as values") to lock it in.


Created .exe from the .py file using PyInstaller in cmd.exe.

