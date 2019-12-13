from gspread import authorize 
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from calendar import Calendar
from pprint import pprint # ~


def get_title(month, year):
	return months[int(month)].title() + ' ' + year + ' года'

def get_orders_row(workers_orders):
	orders_list = []
	for key in workers_orders.keys():
		orders_list.append(key)
	return orders_list

def get_workers_coll(hours_k_dictionary):
	workers_list = []
	for key in hours_k_dictionary.keys():
		workers_list.append(key)
	return workers_list

def get_orders_row(workers_orders):
	orders_list = []
	for hours_list in workers_orders.values():
		hours_list.pop(1)
		hours_list.pop()
		orders_list.append(hours_list)
	orders_list.sort()
	return orders_list

def get_hours_row(hours_k_dictionary):
	hours_list = []
	for hour in hours_k_dictionary.values():
		hours_list.append(hour)
	return hours_list

def write_title(month, year):
	title = get_title(month, year)
	print('Заполняю заголовок: ' + title)
	worksheet.update_acell('B5', title)
	return 1

def write_workers(workers_list):
	print('Заполняю колонку сотрудников: ')
	row, coll = 7, 2
	for i in range(len(workers_list)):
		worksheet.update_cell(row, coll, workers_list[i])
		row = row + 1
	worksheet.update_cell(row, coll, 'Итого:')
	last_row = row
	return last_row

def write_orders(workers_orders, hours_k_dictionary):
	print('Заполняю поле заказов')
	row, coll = 6, 2
	worksheet.update_cell(row, coll, 'Заказ №')
	coll = coll + 1
	orders = get_orders_row(workers_orders)
	hours = get_hours_row(hours_k_dictionary)
	last_row = write_workers(get_workers_coll(hours_k_dictionary))
	row, coll = 7, 3
	for i in range(len(orders)+3):
		worksheet.update_cell(last_row, coll, '=SUM('+chr(64 + coll) + str(row) + ':' + chr(64 + coll) + str(last_row - 1)+ ')')
		coll = coll + 1
	row, coll = 6, 3
	for order in orders:
		worksheet.update_cell(row, coll, order.pop(0))
		coll = coll + 1

	for word in ['Отпуск', 'Простои', 'Итого:']:
		worksheet.update_cell(row, coll, word)
		coll = coll + 1

	row, coll = 7, 3
	for hour_row in hours:
		for hour in hour_row:
			worksheet.update_cell(row, coll, hour)
			coll = coll + 1
		coll = coll + 2
		worksheet.update_cell(row, coll, '=SUM(C7'+':'+chr(65 + coll - 2) + str(row)+ ')')
		coll = 3
		row = row + 1	
	return 1





workers_orders = {
 'АСИ': ['206', 'Открыт', '0', '0', '0', '0', '0', '0'],
 'МБ"Агат"': ['202', 'Открыт', '0', '0', '0', '0', '0', '0'],
 'Методработа': ['299', 'Открыт', '45', '45', '45', '45', '45', '45'],
 'Мини-карт': ['208', 'Приостановлен', '5', '5', '5', '5', '5', '5'],
 'Минитрактор': ['204', 'Открыт', '0', '0', '0', '0', '0', '0'],
 'Самолет': ['207', 'Открыт', '0', '0', '0', '0', '0', '0'],
 'Станок': ['205', 'Приостановлен', '0', '0', '0', '0', '0', '0'],
 'УБА': ['201', 'Приостановлен', '0', '0', '0', '0', '0', '0'],
 'Управление': ['200', 'Открыт', '50', '50', '50', '50', '50', '50']
}


hours_k_dictionary = {
 'Гиловян В.Р.': [43, 39, 0, 0, 0, 0, 0, 0, 4],
 'Данилишин Д.Д.': [43, 39, 0, 0, 0, 0, 0, 0, 4],
 'Данилишин Н.Д.': [43, 39, 0, 0, 0, 0, 0, 0, 4],
 'Литвинн А.В.': [25, 22, 0, 0, 0, 0, 0, 0, 2],
 'Шмаёв М.Ю.': [87, 78, 0, 0, 0, 0, 0, 0, 8]
}

hours_staff = {
 'Гиловян В.Р.': 87,
 'Данилишин Д.Д.': 87,
 'Данилишин Н.Д.': 87,
 'Литвинн А.В.': 51,
 'Шмаёв М.Ю.': 175
}


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
client = authorize(creds)
sh = client.open("LaborReportHours")  # Open the spreadhseet
worksheet = sh.worksheet("ReportHours")
month = datetime.today().strftime('%m' )
year = datetime.today().strftime('%Y' )

write_title(month, year)
write_orders(workers_orders, hours_k_dictionary)