
Applies to fix_cpf_export_format.py:
Create .exe from the .py file using PyInstaller in cmd.exe.

Copy this .exe to a folder that contains .XLS exports from CPFs.
When double-clicked, the program will transcribe the data from each .XLS export
to one .xlsx file.
Now you can easily reference the data in those exports from another spreadsheet
(like the CPF fault aggregator spreadsheet).

To get the CPF fault aggregator spreadsheet to read values from the other
spreadsheets, you have to have the source file open. Once all the data appears
in the sheet though, you can copy the whole thing and paste it over itself
("paste as values") to lock it in.

