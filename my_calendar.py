from flask import Flask, request
import pandas as pd
from ics import Calendar, Event
from datetime import datetime
from pytz import timezone

app = Flask(__name__)

calendar_route = '/module3_ta_schedule.xlsx'
local_tz = timezone('Europe/Amsterdam')

# get from file on computer
def get_file():
    with open(calendar_route, 'rb') as my_file:
        df = pd.read_excel(my_file, 1)
    return df

# process the data according name and surnmae in request and return the calendar

@app.route('/calendar')
def get_calendar():
    # setup name
    name = request.args.get('name')
    surname = request.args.get('surname')
    # name = "Serkan"
    # surname = "Akin"
    name_surname = name + " " + surname

    # get file
    df = get_file()

    # get colomn of TA and their assigned hours
    assigned_col = df[name_surname]

    # get date, hour and duration, lecture name
    data = df[['Date', 'Start at', 'End at', 'Task']]

    cal = Calendar()
    for i in range(len(data)):
        # check whether TA is assignd or not 
        if assigned_col.iloc[i] != 'nan' and str(assigned_col[i]).find('X') != -1 and str(data.iloc[i]['Date']) != 'NaT' and str(data.iloc[i]['Start at']) != 'nan':
            event = Event()
            event.name = data.iloc[i]['Task']
            start_event = localize_time(data.iloc[i]['Date'], data.iloc[i]['Start at'])
            end_event = localize_time(data.iloc[i]['Date'], data.iloc[i]['End at'])
            event.begin = start_event
            event.end = end_event
            cal.events.add(event)
    # with open('my.ics', 'w') as my_file:
    #     my_file.writelines(cal)
    return str(cal), 200, {'Content-Type': 'text/calendar; charset=utf-8'}


def localize_time(date, time):
    naive_datetime = datetime.combine(date, time)
    local_datetime = local_tz.localize(naive_datetime)
    return local_datetime

if __name__ == '__main__':
    app.run(debug=True)


