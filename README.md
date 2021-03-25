# Simple Data Combiner
This project is to pick up the “Work in progress” script, and add functionality to an already existing program.

The following improvements were added: 

    - Select input & output folder with GUI, instead of hardcoding 

    - Remove empty columns when processing data 

    - Provide a combined single-sheet output, with extra column which will show where the data comes from (currently it’s one sheet = one sheet name = one asset type) 

    - Provide a combined csv file


This program has a simple GUI developed using PySimpleGUI. This executable locates in the subfolder named "dist" in the workspace. You can go into the dist folder to find dbf_combiner.exe and run it. You will need to select the workspace folder as an input and the output will be a subfolder named 'excel' containing the combined excel sheet and csv file.

PyInstaller was used to convert Python code into an executable for Windows. 
```
pyinstaller --onefile --noconsole --hidden-import cmath dbf_combiner.py
```

