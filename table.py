import PySimpleGUI as sg


def create_table(headers, data):
    print(headers)
    print(data)

    event_data_window_layout = [
        [sg.Table(
            values=data,
            headings=headers,
            max_col_width=35,
            auto_size_columns=True,
            justification='right',
            num_rows=10,
            key='-TABLE-',
            row_height=35,
            tooltip="Table"
        )]
    ]

    event_data_window = sg.Window("Events", event_data_window_layout, modal=True)

    while True:
        event, values = event_data_window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break

    event_data_window.close()
