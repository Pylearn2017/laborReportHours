import gspread
import locale 
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import datetime
import calendar
import time

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

def get_weekends_list(month, year):
	weekends_list = []
	month = int(month)
	year = int(year)
	for week in cal.monthdayscalendar(year, month):
		weekends_list += (week[-2:])
	return weekends_list

def get_workers_dictionary():
	workers = {}
	first_worker_name = worksheet.find('Штат')
	row = int(first_worker_name.row) + 1
	worker_name = worksheet.acell('A' + str(row)).value

	while worker_name:
		row += 1
		new_worker_name = worksheet.acell('A' + str(row)).value
		worker_name = new_worker_name
		workers[worker_name] = worksheet.acell('B' + str(row)).value
	workers = {k: v for k, v in workers.items() if v}
	return workers	

def get_vacation_days_list():
	vacation_days_list = []
	workers = get_workers_dictionary()

	first_vacation_period = worksheet.find('Отпуска')
	row = int(first_vacation_period.row) + 1
	vacation_day = worksheet.acell('A' + str(row)).value

	while vacation_day:
		row += 1
		new_vacation_day = worksheet.acell('A' + str(row)).value
		vacation_day = new_vacation_day
		list_vacation_day = worksheet.row_values(row)
		vacation_days_list.append(list(filter(None, list_vacation_day)))
	vacation_days_list = list(filter(None, vacation_days_list))
	return vacation_days_list	

def get_holydays_list(weekends_list, m, y):
	month_num = int(m)
	month = months[month_num]
	holydays = []
	holydays_list = []
	first_holyday = worksheet.find('Праздники')
	row = int(first_holyday.row) + 1
	holyday = worksheet.acell('A' + str(row)).value

	while holyday:
		row += 1
		new_holyday = worksheet.acell('A' + str(row)).value
		holyday = new_holyday
		list_holydays = worksheet.row_values(row)
		holydays.append(list(filter(None, list_holydays)))
	for holyday_month in holydays:
		try:
			if month == holyday_month[1]:
				holydays_list += holyday_month[2:]
		except:
			continue		
	return holydays_list	

def check_half_day():
	holydays = get_holydays_list()
	for index in holydays:
		if index:
			if index[1] == month:
				if '*' in row:
					return work_hours - 1
				else:
					return work_hours	

def get_hours_staff(work_days, workers_dictionary, vacation_days_list, weekends_list, holydays_list, m, y):
	month_num = int(m)
	month = months[month_num]
	hours_staff_dictionary = {}
	vacation_staff_dictionary = {}
	work_days_staff_dictionary = {}
	work_hours_staff_dictionary = {}
	vacation_period_without_weekends = {}
	for vacation_days in vacation_days_list:
		if vacation_days[1] == month:
			split_list = [vacation_days[2:][i:i+2] for i in range(0, len(vacation_days[2:])-1)]
			for vacation_periods in split_list:
				vacation_period = set(range(int(vacation_periods[0]), int(vacation_periods[1])+1))
				weekends_list = set(weekends_list)
				vacation_period_without_weekends = set.difference_update(vacation_period, weekends_list)
				vacation_staff_dictionary[vacation_days[0]] = len(vacation_period)
	if '*' in holydays_list:
		work_days = work_days - len(holydays_list) + 1
	else:
		work_days = work_days - len(holydays_list)

	for name, v in workers_dictionary.items():

		hours_staff = int(v) / 5
		hours_staff_dictionary[name] = hours_staff

		try:
			work_days_staff_dictionary[name] = work_days - vacation_staff_dictionary[name]
		except:
			work_days_staff_dictionary[name] = work_days	
		
		work_hours_staff_dictionary[name] = int(hours_staff_dictionary[name] * work_days_staff_dictionary[name]) 
		if '*' in holydays_list:
			work_hours_staff_dictionary[name] -= 1
	print(work_days)
	return work_hours_staff_dictionary


months = ['',
	'январь',
	'февраль',
	'март',
	'апрель',
	'май',
	'июнь',
	'июль',
	'август',
	'сентябрь',
	'октябрь',
	'ноябрь',
	'декабрь',
]

scope = [
	"https://spreadsheets.google.com/feeds",
	'https://www.googleapis.com/auth/spreadsheets',
	"https://www.googleapis.com/auth/drive.file",
	"https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

sh = client.open("LaborReportHours")  # Open the spreadhseet

worksheet = sh.worksheet("Settings")



m = datetime.datetime.today().strftime('%m' )
y = datetime.datetime.today().strftime('%Y' )
m = 3

weekends_list = get_weekends_list(m, y)
work_days = get_wokr_day_without_holiday(m, y)
workers_dictionary = get_workers_dictionary()
vacation_days_list = get_vacation_days_list()
holydays_list = get_holydays_list(weekends_list, m, y)

pprint(get_hours_staff(work_days,workers_dictionary, vacation_days_list, weekends_list, holydays_list, m, y))	


#title = months[int(m)-1] + ' ' + y + ' года'
#worksheet.update_acell('B5', title)
