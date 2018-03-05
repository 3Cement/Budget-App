import os, xlrd, operator, requests, sys, webbrowser, bs4, time
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.compat import range
from tabulate import tabulate
from operator import itemgetter, attrgetter
#from openpyxl.cell import get_column_letter
main_file = "BudgetFile.xlsx" 
all_expenditures = []
categories = {}
def working_directory_data():
	cwd = os.getcwd()
	print(cwd)
	list_of_files = os.listdir('.')
	print(list_of_files)
def yes_or_no(question):
	reply = input(question+' (y/n):\n').lower().strip()
	if reply[:1] == 'y' or reply == '':
		print('Great.')
		return True
	elif reply[:1] == 'n':
		print('Bad to hear that.')
		return False
	else:
		return yes_or_no('Please type Yes or No ')
def start_program():
	start_again = yes_or_no('Would you like to start again?')
	if start_again == True:
		main()
	else:
		print('We are sorry to hear that.')
def test_excel(user_name):
	filename = create_user_workbook(user_name)
	test_expenditures(filename)
	open_excel_file(filename)
def open_file(filename):
	try:
		print('Oppening main file...')
		os.startfile(filename)
	except Exception as exc:
		print('There was a problem: %s' % (exc))
def id_update(active_user):
	wb = load_workbook(main_file)
	ws = wb[active_user]
	for row_num in range(2, ws.max_row+1):
		ws.cell(row=row_num,column=1).value = row_num-1
	wb.save(main_file)

class User():
	_registry = []
	class_counter = 1
	def __init__(self,name,email,password):
		self._registry.append(self)
		self.id = User.class_counter
		self.name = name
		self.email = email
		self.password = password
		print('User {}.{} created.'.format(self.id,self.name))
		User.class_counter += 1
def budget_file():
	if os.path.isfile(main_file) == True:
		wb = load_workbook(filename=main_file)
		print(main_file+" exist. File opened!")
	else:
		wb = Workbook()
		filename = main_file
		ws1 = wb.active
		ws1.title = "Users"
		rows = ["Name", "Email", "Password"]
		ws1.append(rows)
		ws1.freeze_panes = "A2"
		wb.save(filename=filename)
		print(main_file+" created. File opened!")
		return filename
def create_user():
	print('Creating new user...')
	name = input('Please add user name: ')
	email = input('Please add user email: ')
	password = input('Please add user password: ')
	new_user = User(name,email,password)
	wb = load_workbook(filename=main_file)
	ws = wb["Users"]
	#ws = wb.active
	ws.append([name,email,password])
	ws2 = wb.create_sheet(title=name)
	expenditure = ["id", "title", "category", "m_category", "price", "date"]
	ws2.append(expenditure)
	wb.save(main_file)
	print('User '+name+' created!')
	return new_user
def user_names():
	wb = load_workbook(filename=main_file)
	ws = wb["Users"]
	user_names_list = []
	for rowNum in range(2, ws.max_row+1):
		user_name = ws.cell(row=rowNum, column=1).value
		user_names_list.append(user_name)
	print(user_names_list)
	return user_names_list
def user_exist():
	user_names_list = user_names()
	checking_exist = 3
	while checking_exist != 0:
		checking_exist -= 1
		name = input('Please type your account name:\n')
		if name in user_names_list:
			print('Great {}! We found your account!'.format(name))
			return name	
		else:
			print('Wrong username!')
	else:
		print('There is no user with that name. Start again.')
		return False
def show_user_data(name):
	#user_name = input('Please type user name: ')
	wb = load_workbook(filename=main_file)
	worksheet = wb.get_sheet_by_name("Users")
	for rowNum in range(2, worksheet.max_row+1):
		if name == worksheet.cell(row=rowNum, column=1).value:
			print("Great, user "+user_name+" founded!")
			email = worksheet.cell(row=rowNum, column=2).value
			password = worksheet.cell(row=rowNum, column=3).value
			print('User email: '+str(email)+', password: '+str(password))
def user_login(name):
	logging_in = 3
	print('Hello {}! Please log in.'.format(name))
	while logging_in != 0:
		logging_in -= 1
		email = input('Please type your email adress:\n')
		password = input('Please type your password:\n')
		wb = load_workbook(filename=main_file)
		worksheet = wb["Users"]
		for rowNum in range(2, worksheet.max_row+1):
			wb_name = worksheet.cell(row=rowNum, column=1).value
			wb_email = worksheet.cell(row=rowNum, column=2).value
			wb_password = worksheet.cell(row=rowNum, column=3).value
			if wb_name == name and wb_email == email and wb_password == password:
				print('Hello {}! You logged in!'.format(name))
				#break
				ws = wb[name]
				return name
		else:
			print('Wrong email or password!')
	else:
		print('Something go wrong with logging in. Start again')
		return False
def active_user():
	active_user = input("Please put your user name:")
	return active_user

class Expenditure():
	""" This is expenditure """
	_registry = []
	class_counter = 1
	def __init__(self,title,category,m_category,price,date):
		self._registry.append(self)
		self.id = Expenditure.class_counter
		self.title = title
		self.category = category
		self.m_category = m_category
		self.price = price
		self.date = date
		print('Expenditure %s created.' % (title))
		Expenditure.class_counter += 1
	def __repr__(self):
		return('\nExpense nr %s\nTitle: %s\nCategory: %s\nMain category: %s\nPrice: %s$\nDate: %s' % 
			(self.id, self.title, self.category, self.m_category, self.price, self.date))
	def add_to_excel(self,active_user):
		wb = load_workbook(filename=main_file)
		ws = wb[active_user]
		print('Active user is: '+str(active_user))
		to_excel = [self.id,self.title,self.category,self.m_category,self.price,self.date]
		ws.append(to_excel)
		print(self.title+' added to excel file!')
		wb.save(filename=main_file)
def create_expenditure(active_user):
	title = input('Please add expense title: ')
	category = input('Please add expense category: ')
	m_category = input('Please add expense main category: ')
	price = float(input('Please add price: '))
	date = input('Please add date: ')
	new_expenditure = Expenditure(title,category,m_category,-price,date)
	new_expenditure.add_to_excel(active_user)
	all_expenditures.append(new_expenditure)
def create_expenditure2(active_user,title,category,m_category,price,date):
	new_expenditure = Expenditure(title,category,m_category,-price,date)
	new_expenditure.add_to_excel(active_user)
	all_expenditures.append(new_expenditure)
def test_expenditures(active_user):
	create_expenditure2(active_user,"Piwko", "Spożywka", "Jedzenie", 2.35, "25 luty")
	create_expenditure2(active_user,"Prąd", "Opłaty", "Mieszkanie", 6.5, "2017-12-01")
	create_expenditure2(active_user,"Czynsz", "Meble", "Mieszkanie", 123, "10 grudnia")
	create_expenditure2(active_user,"Bilet", "Autobus", "Transport", 2.5, "12 stycznia")
	create_expenditure2(active_user,"Opłata za taxi", "Taxi", "Transport", 2.5, "12 stycznia")
	create_expenditure2(active_user,"Szynka", "Mięso", "Jedzenie", 2.5, "27 marca")
	create_expenditure2(active_user,"Bilet do kina", "Kino", "Rozrywka", 2.5, "31 stycznia")
def delete_expenditure(active_user,wb):
	reply = input("Select the expenditure to be removed:\n")
	old_sheet = wb[active_user]
	old_sheet.title = active_user+'_old'
	max_row = old_sheet.max_row
	max_col = old_sheet.max_column
	wb.create_sheet(active_user)
	new_sheet = wb[active_user]
	# Do the header.
	for col_num in range(1, max_col+1):
		new_sheet.cell(row=1, column=col_num).value = old_sheet.cell(row=1, column=col_num).value
	for row_num in range(2, max_row):
		if reply == str(row_num):
			print('Expenditure '+old_sheet.cell(row=(int(reply)+1), column=2).value+' deleted')
			for row_num in range(1, int(reply)+1):
				for col_num in range(1, max_col+1):
					new_sheet.cell(row=row_num,column=col_num).value=old_sheet.cell(row=row_num,column=col_num).value
				#print("Before done!")
			for row_num in range(int(reply), max_row):
				for col_num in range(1, max_col+1):
					new_sheet.cell(row=(row_num+1),column=col_num).value=old_sheet.cell(row=(row_num+2),column=col_num).value
				#print("After done!")
	wb.remove(old_sheet)
	wb.save(main_file)
	id_update(active_user)
def show_expenditures(active_user,wb,ws):
	tableData = []
	print("All expenditures:")
	for col in ws.iter_rows(min_row=1, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
		tableData.append([col[0].value,col[1].value,col[2].value,col[3].value,col[4].value,col[5].value])
	print(tabulate((tableData),tablefmt="grid"))
def show_transactions_by(active_user,wb,ws):
	tableData = []
	for col in ws.iter_rows(min_row=1, min_col=1, max_row=ws.max_row, max_col=ws.max_column):
		tableData.append((col[0].value,col[1].value,col[2].value,col[3].value,col[4].value,col[5].value))
	tableData.sort(key=operator.itemgetter(2))
	for elem in tableData:
		print(elem)
	print(tableData)
def show_categories():
	print('List of all categories: ')
	for elem in all_expenditures:
		print(elem.category)
def show_main_categories():
	print('List of all main categories: ')
	for elem in all_expenditures:
		print(elem.m_category)

def start():
	working_directory_data()
	budget_file()
	user_names()
	reply = yes_or_no('Hello. Do you already have user account?')
	if reply == True:
		name = user_exist()
		if name == False:
			start_program()
		else:
			user_login(name)
			print(name)
			test_expenditures(name)
			id_update(name)
			what_next(name)
			#open_file(main_file)
	elif reply == False:
		print('Then you need to create new user: ')
		name = user_login(create_user().name)
		print(name)
		test_expenditures(name)
		what_next(name)
		#open_file(main_file)
	else:
		print('Something go wrong with your name!')
def what_next(active_user):
	wb = load_workbook(main_file)
	ws = wb[active_user]
	print('Please choose what would you like to do next')
	i = 1
	strings = [	"Show all expenditures",
				"Add new expenditure",
				"Delete expenditure",
				"Show all incomes",
				"Add new income",
				"Delete income",
				"Show transactions in correct order",
				"Show balances in correct order",
				"Show charts",
				"Change user",
				"Restart program",
				"Save and exit"]
	for string in strings:
		print(string.ljust(35,'-')+str(i).rjust(2,'-'))
		i += 1
	reply = input('Type number from 1 to {}'.format(i).center(37,'-')+'\n')
	if reply == "1":
		show_expenditures(active_user,wb,ws)
	elif reply == "2":
		print("Adding new expenditure:")
		create_expenditure(active_user)
	elif reply == "3":
		print("Delete expenditure:")
		show_expenditures(active_user,wb,ws)
		delete_expenditure(active_user,wb)
	elif reply == "4":
		print("Showing all incomes:")
	elif reply == "5":
		print("Adding new income:")
	elif reply == "6":
		print("Delete income:")
	elif reply == "7":
		print("Showing transaction. Choose order:")
		show_transactions_by(active_user,wb,ws)
	elif reply == "8":
		print("Show balace. Choose order:")
	elif reply == "9":
		print("Show charts. Choose data:")
	elif reply == "10":
		print("Changing user account:")
	elif reply == "11":
		print("Restarting program:")
	elif reply == "12":
		print("Saving data and exit...")
	else:
		print("Error. Something go wrong...")
	open_file(main_file)

start()
