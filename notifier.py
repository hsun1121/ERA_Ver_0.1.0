from pathlib import Path
from plyer import notification
import datetime
import openpyxl

current_dir = Path(__file__).parent
EXCEL_FILE = current_dir / 'Saved_Events.xlsx'


def send_notification():
    today = datetime.datetime.now()
    current_month = today.strftime('%B')
    day = today.strftime('%d')
    year = today.strftime('%Y')

    print(current_month)
    print(day)
    print(year)

    event_data = openpyxl.load_workbook('Saved_Events.xlsx')

    user_data = event_data['Sheet1']

    # Scan rows starting from row 2 in excel file
    for x in range(2,user_data.max_row+1):
        # Check if value in each row and column C (index of 2 will change to 3 for Day and 4 for Year)
        # in excel file is equal to current month
        if str(user_data[x][2].value) == current_month and str(user_data[x][3].value) == day and str(
                user_data[x][4].value) == year:
            # Print name of corresponding event
            # print(str(user_data[x][0].value))

            title = 'Event Reminder'
            message = 'You have an event today: ' + str(user_data[x][0].value)

            notification.notify(title=title,
                                message=message,
                                app_icon=None,
                                timeout=10,
                                toast=False)
