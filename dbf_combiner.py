import PySimpleGUI as sg
import os.path
import combine_data

# First the window layout in 2 columns

file_list_column = [
    [
        sg.Text("Please select input folder"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(),
    ],
    [
        sg.Listbox(
            values=[], enable_events=False, size=(40, 20), key="-FILE LIST-"
        ) # False enable_events to show list of files only
    ],
]

# For now will only show the name of the file that was chosen
submission_column = [
    [sg.Text("This is the folder path you selected:")],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.Button("Ok", key='-SUBMIT-')],
    [sg.Text(size=(40, 1), key="-MESSAGE-")],
    [sg.Button("Exit", size=(10, 1))],
    # [sg.Image(key="-IMAGE-")],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(submission_column),
    ]
]

window = sg.Window("DBF to Combined Excel", layout)

# Run the Event Loop
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED: # keep running until exit or close window
        break
    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith((".dbf"))
        ]
        window["-FILE LIST-"].update(fnames)
        window["-TOUT-"].update(folder)

    elif event == "-SUBMIT-":  # A file was chosen from the listbox
        try:
            folder = os.path.abspath(values["-FOLDER-"])
            combine_data.dbf_to_excel(path=folder)
            message = 'Excel sheet produced successfully!'
            window["-MESSAGE-"].update(message)

        except:
            message = 'Something went wrong.'
            window["-MESSAGE-"].update(message)


window.close()