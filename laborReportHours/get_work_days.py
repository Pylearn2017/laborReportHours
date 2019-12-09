import calendar
from pprint import pprint
from locale import setlocale,LC_ALL
import holydays
import datetime
 
setlocale(LC_ALL, "")

cal = calendar.Calendar()

def get_wokr_day_without_holiday(month, year):
	month = int(month)
	year = int(year)
	weekday_count = 0
	for week in cal.monthdayscalendar(year, month):
		for i, day in enumerate(week):
			if day == 0 or i >= 5:
				continue
			weekday_count += 1
	return weekday_count

def get_work_days(month, year, vacation_days = 0):
	month = int(month)
	year = int(year)
	holydays_list = []
	for index in holydays.holydays.items():
		if index[1][1] == month:
			holydays_list += index[1][0]
	for week in cal.monthdayscalendar(year, month):
		for holyday in holydays_list:
			if holyday in week:
				vacation_days += 1		
	work_days = get_wokr_day_without_holiday(month, year) - vacation_days
	return work_days

	

m = datetime.datetime.today().strftime('%m' )
y = datetime.datetime.today().strftime('%Y' )
