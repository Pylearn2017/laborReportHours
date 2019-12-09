import calendar
import datetime

cal = calendar.Calendar()

def get_weekends(m, y):
	weekends_list = []
	month = int(m)
	year = int(y)
	for week in cal.monthdayscalendar(year, month):
		print(week[-2:])
		weekends_list += (week[-2:])
	return weekends_list	

m = datetime.datetime.today().strftime('%m' )
y = datetime.datetime.today().strftime('%Y' )

print(get_weekends(m,y))