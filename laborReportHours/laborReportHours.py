from gspread import authorize 
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from calendar import Calendar
from pprint import pprint # ~

cal = Calendar()

def get_wokr_day_without_holiday(month, year):
	month = int(month)
	year = int(year)
	print('Получаю количество дней месяца', months[month], year, 'года', end='...')
	weekday_count = 0
	for week in cal.monthdayscalendar(year, month):
		for i, day in enumerate(week):
			if day == 0 or i >= 5:
				continue
			weekday_count += 1
	print('✔')
	return weekday_count

def get_weekends_list(month, year):
	weekends_list = []
	month = int(month)
	year = int(year)
	for week in cal.monthdayscalendar(year, month):
		weekends_list += (week[-2:])
	return weekends_list

def get_workers_dictionary():
	print('Получаю список работников', end='...')
	workers = {}
	first_worker_name = worksheet.find('Штат')
	row = int(first_worker_name.row) + 1
	worker_name = worksheet.acell('A' + str(row)).value

	while worker_name:
		row += 1
		new_worker_name = worksheet.acell('A' + str(row)).value
		worker_name = new_worker_name
		workers[worker_name] = worksheet.acell('B' + str(row)).value
		print('.', end='.')
	workers = {k: v for k, v in workers.items() if v}
	print('✔')	
	return workers	

def get_workers_orders():
	print('Получаю список заказов', end='...')
	orders = {}
	first_order = worksheet.find('Заказы')
	row = int(first_order.row) + 1
	order = worksheet.acell('A' + str(row)).value
	k_list = []

	while order:
		row += 1
		new_order = worksheet.acell('A' + str(row)).value
		order = new_order
		if worksheet.acell('C' + str(row)).value in ['Открыт', 'Приостановлен'] :
			list_order_k = worksheet.row_values(row)
			orders[order] = list(filter(None, list_order_k[1:]))
			print('.', end='.')
	orders = {k: v for k, v in orders.items() if v}
	print('✔')	
	return orders


def get_vacation_days_list(workers_dictionary):
	print('Получаю диапазоны отпусков', end='...')
	vacation_days_list = []
	workers = workers_dictionary

	first_vacation_period = worksheet.find('Отпуска')
	row = int(first_vacation_period.row) + 1
	vacation_day = worksheet.acell('A' + str(row)).value

	while vacation_day:
		row += 1
		new_vacation_day = worksheet.acell('A' + str(row)).value
		vacation_day = new_vacation_day
		list_vacation_day = worksheet.row_values(row)
		vacation_days_list.append(list(filter(None, list_vacation_day)))
		print('.', end='.')
	vacation_days_list = list(filter(None, vacation_days_list))
	print('✔')	
	return vacation_days_list	

def get_holydays_list(weekends_list, m, y):
	print('Получаю список праздничных дней', end='...')
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
		print('.', end='.')
	for holyday_month in holydays:
		try:
			if month == holyday_month[1]:
				holydays_list += holyday_month[2:]
		except:
			continue	
	print('✔')				
	return holydays_list	

def check_half_day():
	print('Проверяю * - короткие дни', end='...')
	holydays = get_holydays_list()
	for index in holydays:
		if index:
			if index[1] == month:
				if '*' in row:
					print('✔')	
					return work_hours - 1

				else:
					print('✔')	
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
	print('Считаю часы для', work_days, 'рабочих дней....✔')
	return work_hours_staff_dictionary

def get_k_dictionary(workers_dictionary, workers_orders):
	k_dictionary = {}
	new_m = []
	i = 0
	m = list(workers_orders.values())
	for item in m:
		new_m.append(item[2:])
	matrix = list(zip(*new_m))
	for worker in workers_dictionary.keys():
		k_dictionary[worker] = matrix[i]
		i +=  1	
	return k_dictionary		

def get_row_orders(workers_orders):
	row_orders = []
	for order, hours in workers_orders.items():
		row_orders.append(order + ' - ' + hours[0])
	row_orders.append('Отпуск')
	row_orders.append('Простои')
	return row_orders	

def get_hours_k_dictionary(k_dictionary, hours_staff):
	hours_k_dictionary = {}
	k_hours = []
	for worker, hours in hours_staff.items():
		for k in k_dictionary[worker]:
			k_hour = int(k) * hours
			k_hours.append(int(k_hour/100))
		hours_k_dictionary[worker] = k_hours
		k_hours = []	
	return hours_k_dictionary

def write_title(month, year):
	pass




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
try:
	creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
except:
	print('Нет доступа, обратитесь к разработчику alexlik73@gmail.com')
finally:
	pass
	#input('Для продолжения нажмите Enter')		

client = authorize(creds)

sh = client.open("LaborReportHours")  # Open the spreadhseet

worksheet = sh.worksheet("Settings")

m = datetime.today().strftime('%m' )
y = datetime.today().strftime('%Y' )

weekends_list = get_weekends_list(m, y)
work_days = get_wokr_day_without_holiday(m, y)
workers_dictionary = get_workers_dictionary()
workers_orders = get_workers_orders()
vacation_days_list = get_vacation_days_list(workers_dictionary)
holydays_list = get_holydays_list(weekends_list, m, y)
hours_staff = get_hours_staff(work_days,workers_dictionary, vacation_days_list, weekends_list, holydays_list, m, y)
k_dictionary = get_k_dictionary(hours_staff, workers_orders)
hours_k_dictionary = get_hours_k_dictionary(k_dictionary, hours_staff)
#row_orders = get_row_orders(workers_orders)
pprint(hours_k_dictionary)
pprint(workers_orders)
pprint(hours_staff)
#title = months[int(m)-1] + ' ' + y + ' года'
#worksheet.update_acell('B5', title)

