from pathlib import Path

import PySimpleGUI as sg
import pandas as pd
import numpy as np

from table import create_table
from enter_event import enter_data

# Add some color to the window
sg.theme('DarkTeal9')

# Variables
current_dir = Path(__file__).parent
EXCEL_FILE = current_dir / 'Saved_Events.xlsx'

# Home GUI window
layout = [
    [sg.Text('What would you like to do:')],

    [sg.Button('Add Event'), sg.Button('Show Events'), sg.Exit()]
]

window = sg.Window('Event Reminder App', layout)


# Control for home GUI window
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Add Event':
        enter_data(EXCEL_FILE)
    if event == 'Show Events':
        excel_file_df = pd.read_excel(EXCEL_FILE)
        headers = excel_file_df.columns.to_numpy()
        data_array = excel_file_df.to_numpy()
        # Create ndarray and past to table.py
        create_table(headers.tolist(), data_array.tolist())

window.close()


