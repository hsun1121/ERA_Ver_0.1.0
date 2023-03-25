from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
import numpy as np
from enter_event import enter_data
from event_table import create_table

current_dir = Path(__file__).parent
EXCEL_FILE = current_dir / 'Saved_Events.xlsx'


def run_home_gui():
    # Home GUI window
    layout = [
        [sg.Text('What would you like to do:')],

        [sg.Button('Add Event'), sg.Button('Show Events'), sg.Exit()]
    ]

    window = sg.Window('Event Reminder App', layout)
    home_gui_control(window)
    return window


def home_gui_control(window):
    # Control for home GUI window
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Add Event':
            enter_data()
        if event == 'Show Events':
            excel_file_df = pd.read_excel(EXCEL_FILE)
            headers = excel_file_df.columns.to_numpy()
            # Create ndarray to paste to event_table.py
            data_array = excel_file_df.to_numpy()
            # Create dataframe to replace NaN with blanks
            data_array2 = pd.DataFrame(data_array)
            # After replacing NaN's, turn dataframe back to numpy array
            final_data_array = data_array2.replace(np.nan, '').to_numpy()
            # List out all values from excel file onto GUI
            create_table(headers.tolist(), final_data_array.tolist())

    window.close()