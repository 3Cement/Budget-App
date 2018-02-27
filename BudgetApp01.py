import os, xlrd, operator
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.compat import range
from openpyxl.cell import get_column_letter
def working_directory_data():
	cwd = os.getcwd()
	print(cwd)
	list_of_files = os.listdir('.')
	print(list_of_files)
class User():
	_registry = []
	class_counter = 1
	def __init__(self,name,email,password):
		self._registry.append(self)
		self.id = User.class_counter
		self.name = name
		self.email = email
		self.password = password
		print('User %s.%s created.' % (self.id,self.name))
		User.class_counter += 1
def create_user_workbook(user_name):
	expenditure = ["id", "title", "category", "m_category", "price", "date"]
	wb = Workbook()
	filename = user_name+'_budget.xlsx'
	ws1 = wb.active
	ws1.title = "Expenditures"
	ws1.append(expenditure)
	print("Workbook "+filename+" created!")
	wb.save(filename=filename)
	return filename
def create_user():
	print('Creating new user...')
	name = input('Please add user name: ')
	email = input('Please add user email: ')
	password = input('Please add user password: ')
	new_user = User(name,email,password)
	create_user_workbook(name)
	return new_user
def show_users():
	user_names =[]
	#print('All '+str(User.class_counter-1)+' users below:')
	for user in User._registry:
		print(str(user.id)+'. '+user.name)
		user_names.append(user.name)
	return user_names
def show_users_data():
	show_users()
	print('All '+str(User.class_counter-1)+' users below:')
	for user in User._registry:
		print(str(user.id)+'. '+user.name+' email: '+user.email+' password: '+user.password)
		user_names.append(user.name)
def user_exist():
	user_names =[]
	for user in User._registry:
		user_names.append(user.name)
	print(user_names)
	checking_exist = 3
	while checking_exist != 0:
		checking_exist -= 1
		name = input('Please type your account name:\n')
		if name in user_names:
			print('Great %s! We found your account!' % (name))
			return name	
		else:
			print('Wrong username!')
	else:
		print('There is no user with that name. Start again.')
		return False
def user_login(name):
	logging_in = 3
	print('Hello %s! Please log in.' % (name))
	while logging_in != 0:
		logging_in -= 1
		email = input('Please type your email adress:\n')
		password = input('Please type your password:\n')
		for user in User._registry:
			if user.name == name and user.email == email and user.password == password:
				print('Hello %s! You logged in!' % (user.name))
				#break
				return True
		else:
			print('Wrong email or password!')
	else:
		print('Something go wrong with logging in. Start again')
		return False
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

all_expenditures = []
categories = {}

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

def create_expenditure():
		title = input('Please add expense title: ')
		category = input('Please add expense category: ')
		m_category = input('Please add expense main category: ')
		price = float(input('Please add price: '))
		date = input('Please add date: ')
		new_expenditure = Expenditure(title,category,m_category,price,date)
		all_expenditures.append(new_expenditure)
		return new_expenditure

def total_expenses():
	total_expenses = 0
	for expense in Expenditure._registry:
		total_expenses += expense.price
	print('Total expenses: ' + str(total_expenses) +'$')

exp1 = Expenditure("Piwko", "Spożywka", "Jedzenie", 2.35, "25 luty")
all_expenditures.append(exp1)
exp2 = Expenditure("Prąd", "Opłaty", "Mieszkanie", 6.5, "2017-12-01")
all_expenditures.append(exp2)
exp3 = Expenditure("Czynsz", "Meble", "Mieszkanie", 123, "10 grudnia")
all_expenditures.append(exp3)
exp4 = Expenditure("Bilet", "Autobus", "Transport", 2.5, "12 stycznia")
all_expenditures.append(exp4)
exp5 = Expenditure("Opłata za taxi", "Taxi", "Transport", 2.5, "12 stycznia")
all_expenditures.append(exp5)
exp6 = Expenditure("Szynka", "Mięso", "Jedzenie", 2.5, "12 stycznia")
all_expenditures.append(exp6)
exp7 = Expenditure("Bilet do kina", "Kino", "Rozrywka", 2.5, "12 stycznia")
all_expenditures.append(exp7)

def show_categories():
	print('List of all categories: ')
	for elem in all_expenditures:
		print(elem.category)

def show_main_categories():
	print('List of all main categories: ')
	for elem in all_expenditures:
		print(elem.m_category)

def expenditures_sorted_by(name):
	#name = input('Select the sort key: id, title, category, m_category, price, date:\n')
	print('Expenditures sorted by: '+name)
	sorted_all_expenditures = sorted(all_expenditures, key=operator.attrgetter(name))
	for elem in sorted_all_expenditures:
		print(elem)

user1 = User('Daniel', 'd1', 'd2')
user2 = User('Kamil', 'k1', 'k2')
print(user1.name)
show_categories()
show_main_categories()
expenditures_sorted_by("price")

def main():
	reply = yes_or_no('Hello. Do you already have user account?')
	if reply == True:
		name = user_exist()
		if name == False:
			start_program()
		else:
			user_login(name)

	elif reply == False:
		print('Then you need to create new user: ')
		user_login(create_user().name)
	else:
		print('Something go wrong with your name!')
#main()

