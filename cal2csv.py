import sys
import os.path
from icalendar import Calendar
import csv
from datetime import datetime, timedelta, date

filename = sys.argv[1]
filename_noext = filename[:-4]
file_extension = str(sys.argv[1])[-3:]
headers = ('Summary', 'Start Time', 'End Time', 'Hours')
if len(sys.argv) < 3:
    print("No month provided, proceeding with current month")
    month = date.today().month
else:
    month = sys.argv[2]

class CalendarEvent:
    """Calendar event class"""
    summary = ''
    start = ''
    end = ''
    hours = ''

    def __init__(self, name):
        self.name = name

weeks = []
events = []


def open_cal():
    if os.path.isfile(filename):
        if file_extension == 'ics':
            print("Extracting events from file:", filename, "\n")
            f = open(sys.argv[1], 'rb')
            gcal = Calendar.from_ical(f.read())
            for component in gcal.walk():
                event = CalendarEvent("event")
                if component.get('STATUS') != "CONFIRMED":
                    continue
                if component.get('TRANSP') == 'TRANSPARENT' or component.get('TRANSP') == None:
                    continue #skip event that have not been accepted
                if component.get('SUMMARY') == None: continue #skip blank items
                event.summary = component.get('SUMMARY')
                if hasattr(component.get('dtstart'), 'dt'):
                    event.start = component.get('dtstart').dt
                if hasattr(component.get('dtend'), 'dt'):
                    event.end = component.get('dtend').dt
                now = date.today()
                req_date_from = date(now.year, int(month), 1)
                req_date_to = date(now.year, int(month), 31)
                if event.start.date() < req_date_from or event.start.date() > req_date_to:
                    continue
                event.hours = event.end - event.start
                secs = event.hours.seconds
                minutes = ((secs/60)%60)/60.0
                hours = secs/3600
                event.hours = hours + minutes
                events.append(event)
            f.close()
        else:
            print("You entered ", filename, ". ")
            print(file_extension.upper(), " is not a valid file format. Looking for an ICS file.")
            exit(0)
    else:
        print("I can't find the file ", filename, ".")
        print("Please enter an ics file located in the same folder as this script.")
        exit(0)


def csv_write(icsfile):
    try:
        count = 0
        for week in weeks:
            count = count + 1
            csvfile = "week" + str(count) + ".csv"
            with open(csvfile, 'w') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
                wr.writerow(headers)
                for event in week:
                    values = (event.summary.encode('utf8').decode(), event.start, event.end, event.hours)
                    wr.writerow(values)
    except IOError:
        print("Could not open file! Please close Excel!")
        exit(0)


def sort_by_weekly(events):
    week = []
    end_of_week = None
    for event in events:
        if end_of_week == None:
            end_of_week = event.start - timedelta(days=event.start.weekday()) + timedelta(days=6)
            end_of_week = end_of_week.replace(hour=23, minute=59)
        if event.start > end_of_week:
            weeks.append(week)
            week = []
            end_of_week = None
        else:
            week.append(event)
    weeks.append(week)

open_cal()
sortedevents=sorted(events, key=lambda obj: obj.start) # Needed to sort events. They are not fully chronological in a Google Calendard export ...
sort_by_weekly(sortedevents)
csv_write(filename)

from pyexcel.cookbook import merge_all_to_a_book
import glob

merge_all_to_a_book(glob.glob("./*.csv"), filename_noext+".xlsx")
print("Done, your file: "+filename_noext+".xlsx")

count = 0
for week in weeks:
    count = count + 1
    csvfile = "week" + str(count) + ".csv"
    os.remove(csvfile)