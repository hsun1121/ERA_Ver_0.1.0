import PySimpleGUI as sg
import pandas as pd


# Build GUI windows to enter new event to spreadsheet
def enter_data(excel):
    df = pd.read_excel(excel)

    enter_data_layout = [
        [sg.Text('Please fill out the following fields:')],
        [sg.Text('Event Name', size=(15, 1)), sg.InputText(key='Event Name')],
        [sg.Text('Location', size=(15, 1)), sg.InputText(key='Location')],
        [sg.Text('Month', size=(15, 1)),
         sg.Combo(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
                   'September', 'October', 'November', 'December'], key='Month')],
        [sg.Text('Day', size=(15, 1)),
         sg.Combo(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14',
                   '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26',
                   '27', '28', '29', '30', '31'], key='Day')],
        [sg.Text('Year', size=(15, 1)), sg.Combo(['2023', '2024', '2025'], key='Year')],

        [sg.Submit(), sg.Button('Clear'), sg.Exit()]
    ]

    enter_data_window = sg.Window('Add Event', enter_data_layout, modal=True)

    while True:
        event, values = enter_data_window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Clear':
            for key in values:
                enter_data_window[key]('')
        if event == 'Submit':
            new_record = pd.DataFrame(values, index=[0])
            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(excel, index=False)
            sg.popup('Event Saved!')
            enter_data_window.close()

    enter_data_window.close()