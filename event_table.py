from pathlib import Path
import PySimpleGUI as sg
from update_event import update_event_gui

current_dir = Path(__file__).parent
EXCEL_FILE = current_dir / 'Saved_Events.xlsx'


def create_table(headers, data):
    print(headers)
    print(data)

    event_data_window_layout = [
        [sg.Text('Please click on an event to delete it or click on any detail to update it:')],
        [sg.Table(
            values=data,
            headings=headers,
            max_col_width=35,
            auto_size_columns=True,
            justification='right',
            num_rows=10,
            key='-TABLE-',
            row_height=35,
            tooltip="Table",
            enable_click_events=True
        )]
    ]

    event_data_window = sg.Window("Events", event_data_window_layout, modal=True)
    click_events_control(event_data_window)
    return event_data_window


def click_events_control(event_data_window):
    while True:
        event, values = event_data_window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break
        # If an event is clicked on, present update event gui window
        if '+CLICKED+' in event[1]:
            # Assign event row and column numbers from GUI to variables
            row_number = event[2][0]
            column_number = event[2][1]
            update_event_gui(row_number, column_number)
            event_data_window.close()

    event_data_window.close()
