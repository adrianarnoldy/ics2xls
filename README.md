# ics2xls
Taken from https://github.com/erikcox/ical2csv, I extended his script to cater my needs. Using his calender to csv as the base, I developed an extension to separate the converted schedules into weeks, and added a time limiter (eg from 2021-07-01 to 2021-07-31), and then convert it to XLS for better viewing, creating a different sheet for each weeks. I also created the hours spent on those meetings to complete the report.

This script will take only your accepted invites, or invites where you are not optional to attend (normal invite, must come).

In the future I'd also like to improve this script to make my WFH life easier 🙏.

Dependencies:
* Calendar (pip3 install icalendar)
* PyExcel (pip3 install pyexcel pyexcel-xlsx)

Convert your calendar (.ics) file to an excel (.xls) file using this simple script.

How to use:
```
python3 cal2csv.py <your_ics_file> <month>

example: python3 cal2csv.py testing.ics 7
this will generate your schedules in july
```

You can use this script however you want. 👍
