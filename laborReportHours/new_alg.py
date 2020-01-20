from oauth2client.service_account import ServiceAccountCredentials
from gspread import authorize 
from calendar import Calendar
from math import ceil
from time import sleep

def get_year():
	text = 'Получаю год отчета '
	print(text)
	data_place = worksheet.find('Дата')
	row = int(data_place.row) + 1
	year = worksheet.acell('B' + str(row + 1)).value
	return year	

def get_month(months):
	print('Получаю месяц отчета ')
	data_place = worksheet.find('Дата')
	row = int(data_place.row) + 1
	month = worksheet.acell('A' + str(row + 1)).value
	month = months.index(month)
	if len(str(month)) <= 2:
		month = 0 + month
	return month

def get_days_month(month, year):
	print('Получаю список всех дней')
	month = int(month)
	year = int(year)
	weekday_count = 0
	days_month = cal.monthdayscalendar(year, month)
	return days_month

def get_holydays_list(m):
	print('Получаю список праздничных дней')
	month_num = int(m)
	month = months[month_num]
	holydays = []
	holydays_list = []
	first_holyday = worksheet.find('Праздники')
	row = int(first_holyday.row) + 1
	holyday = worksheet.acell('A' + str(row)).value

	while holyday:
		sleep(1)
		row += 1
		sleep(1)
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

def get_vacation_days_list():
	print('Получаю диапазоны отпусков')
	vacation_days_list = []

	first_vacation_period = worksheet.find('Отпуска')
	row = int(first_vacation_period.row) + 1
	vacation_day = worksheet.acell('A' + str(row)).value

	while vacation_day:
		sleep(1)
		row += 1
		new_vacation_day = worksheet.acell('A' + str(row)).value
		vacation_day = new_vacation_day
		list_vacation_day = worksheet.row_values(row)
		vacation_days_list.append(list(filter(None, list_vacation_day)))
	vacation_days_list = list(filter(None, vacation_days_list))
	return vacation_days_list

def get_days_without_holydays(days, holydays):
	for week in days:
		for day in week:
			if str(day) in holydays:
				week[week.index(day)] = 0
			if str(day) + '*' in holydays:
				week[week.index(day)] = '*'
	return days

def get_workers_dictionary():
	print('Получаю расписание работников')
	workers = []
	first_finde = worksheet.find('Штат')
	row = int(first_finde.row) + 1
	worker = worksheet.acell('A' + str(row)).value

	while True:
		sleep(1)
		row += 1
		new_worker = worksheet.acell('A' + str(row)).value
		if new_worker:
			worker = new_worker
			work_list = worksheet.row_values(row)
			workers.append(list(filter(None, work_list)))
		else: break	
	return workers

def get_work_days(workers, month_without_holidays):
	dictionary_workers = {}
	month_worker = []
	week_worker = [] 
	for worker in workers:
		hours = worker[1]
		for week in month_without_holidays:
			for day in range(7):
				if week[day] == '*':
					week_worker.append(int(hours)-1)
				elif week[day] != 0 and worker[2:-1][day] != '0':
					week_worker.append(int(hours))
				else: 
					week_worker.append(0)	
			month_worker.append(week_worker)	
			week_worker = []		
		dictionary_workers[worker[0]] = month_worker
		month_worker = []	
	return dictionary_workers

def get_sum_work_hours(dictionary_workers):
	sum_work_hours = {}
	sum_hours_list = []
	for k,v in dictionary_workers.items():
		for week in v:
			sum_hours_list.append(sum(week))		
		sum_work_hours[k] = sum(sum_hours_list)
		sum_hours_list = []
	return sum_work_hours	

def get_workers_orders():
	text = 'Получаю список заказов'
	print(text)
	orders = {}
	first_order = worksheet.find('Заказы')
	row = int(first_order.row) + 1
	order = worksheet.acell('A' + str(row)).value
	k_list = []

	while order:
		sleep(1)
		row += 1
		new_order = worksheet.acell('A' + str(row)).value
		order = new_order
		if worksheet.acell('C' + str(row)).value in ['Открыт', 'Приостановлен'] :
			list_order_k = worksheet.row_values(row)
			orders[order] = list(filter(None, list_order_k[1:]))
	orders = {k: v for k, v in orders.items() if v}
	for k, v in orders.items():
		if 'Приостановлен' in v:
			for ho in range(2, len(v)):
				orders[k][ho] = 0		
	return orders	

def calculate_hours(sum_work_hours, workers_orders):
	calculated_hours_dictionary = {}
	list_remainder_of_division = []	
	k_list = []
	i = 2
	for worker, value in sum_work_hours.items():
		for orders in workers_orders.values():
			k = int(orders[i]) * value / 100
			remainder_of_division = k % 1
			k_list.append(int(k))
			list_remainder_of_division.append(remainder_of_division)	
		last_order = k_list.pop(1)
		k_list.append(last_order)
		sum_list_remainder_of_division = sum(list_remainder_of_division)
		while sum_list_remainder_of_division > 0.5:
			sum_list_remainder_of_division = sum_list_remainder_of_division - 1
			k_list[-1] = k_list[-1] + 1
		calculated_hours_dictionary[worker] = k_list
		i = i + 1	
		k_list = []
		list_remainder_of_division = []
		

	return calculated_hours_dictionary

def get_title(month, year):
	return months[month].title() + ' ' + year + ' года'

def get_list_of_orders_for_writer(workers_orders):
	list_of_orders_for_writer = []
	for order in workers_orders.values():
		list_of_orders_for_writer.append(order[0])
	list_of_orders_for_writer.sort()
	return list_of_orders_for_writer

def write_title(month, year):
	title = get_title(month, year)
	worksheet.update_acell('B5', title)
	return 1

def write_orders(list_of_orders_for_writer):
	row, coll = 6, 2
	worksheet.update_cell(row, coll, 'Заказ №')
	coll = coll + 1
	list_of_orders_for_writer = list_of_orders_for_writer + ['Отпуск:', 'Простои:', 'Итого:']
	for order in list_of_orders_for_writer:
		worksheet.update_cell(row, coll, order)
		coll = coll + 1

def write_labor_report_hours(labor_report_hours, sum_work_hours):
	row, coll = 7, 2
	for name, value in labor_report_hours.items():
		value = value + [0, 0] 
		worksheet.update_cell(row, coll, name)
		sum_worker = sum_work_hours[name]	
		for hour in value:
			sleep(1)
			coll = coll + 1
			worksheet.update_cell(row, coll, hour)
		worksheet.update_cell(row, coll + 1, str(sum_worker))
		row = row + 1
		coll = 2
	coll = coll + 1
	sum_coll = list(labor_report_hours.values())
	for k in sum_coll[1] + [0, 0, 0]:
		worksheet.update_cell(row, coll, '=SUM(' + chr(64 + coll) + str(row-1) + ':' + chr(64 + coll) + '7' + ')')
		coll = coll + 1
	

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

cal = Calendar()
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = authorize(creds)
sh = client.open("LaborReportHours")  # Open the spreadhseet
worksheet = sh.worksheet("Settings")
month = get_month(months)
year = get_year()
days_month = get_days_month(month, year)

holydays = get_holydays_list(month)
# holydays = ['31*']
# print(holydays) 
# print('holydays')

vacation_days = get_vacation_days_list()
# vacation_days = [['Шмаёв М.Ю.', 'декабрь', '1', '3'], ['Данилишин Д.Д.', 'ноябрь'], ['Данилишин Д.Д.', 'декабрь']]
# print(vacation_days)
# print('vacation_days')

month_without_holidays = get_days_without_holydays(days_month, holydays)

workers = get_workers_dictionary()
# workers = [['Гиловян В.Р.', '4', '1', '1', '1', '1', '1', '0', '0', 'K1'], ['Данилишин Д.Д.', '4', '1', '1', '1', '1', '1', '0', '0', 'K2'], ['Данилишин Н.Д.', '4', '1', '1', '1', '1', '1', '0', '0', 'K3'], ['Литвинн А.В.', '4', '1', '1', '0', '0', '1', '0', '0', 'K4'], ['Шмаёв М.Ю.', '8', '1', '1', '1', '1', '1', '0', '0', 'K5']]
# print(workers)
# print('workers')

work_days = get_work_days(workers, month_without_holidays)

workers_orders = get_workers_orders()

sum_work_hours = get_sum_work_hours(work_days)
labor_report_hours = calculate_hours(sum_work_hours, workers_orders)


# Writer
worksheet = sh.worksheet("ReportHours")
list_of_orders_for_writer = get_list_of_orders_for_writer(workers_orders)
write_title(month, year)
write_orders(list_of_orders_for_writer)
write_labor_report_hours(labor_report_hours, sum_work_hours)
#write_vacation_hours(vacation_hours)