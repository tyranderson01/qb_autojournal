import os
from pyautogui import press, typewrite, hotkey, keyDown, keyUp #send keystrokes and hotkeys
import openpyxl
import pyperclip
import time

# qb = os.startfile("C:\\Users\\tyler.anderson\\Desktop\\New QuickBooks.rdp")
# SET THE WORKSHEETS
filepath = "C:\\Users\\tyler.anderson\\Documents\\QB Entry.xlsx"
wb = openpyxl.load_workbook(filepath, data_only=True)
single_ws = wb['Single']['A']
double_ws = wb['Double']
qbref_ws = wb['QB Ref']
gl_ws = wb['GL Detail']
journalws = wb['Journal']

def get_gl_info():
	fac_list = []
	description_list = []
	debit_list = []
	credit_list = []
	index_list = []
	x=0
	for cell in gl_ws['O']:
		fac_list.append(cell.value)
		x=x+1
	
	for i in range(x):
		index_list.append(i)
	
	for cell in gl_ws['H']:
		description_list.append(cell.value)
		
	for cell in gl_ws['L']:
		debit_list.append(cell.value)
	
	for cell in gl_ws['M']:
		credit_list.append(cell.value)
		
	valuelist = list(zip(fac_list,description_list,debit_list,credit_list))
	get_gl_info = dict(zip(index_list,valuelist))
	# print(gl_dict)
	return get_gl_info
		


# COLLECT ALL THE INFO FROM THE WORKSHEETS
def get_single_info():
	data = [] # list to append info to from copyfile
	for cell in single_ws: # collect info into data
		if isinstance(cell.value, float) or isinstance(cell.value, int):
			data.append(round(cell.value,2))
		elif isinstance(cell.value, str):
			data.append(cell.value)
	return data
	

def get_double_info():
	building = [] # list to append info to from copyfile
	amount = []
	descrip = []
	for cell in double_ws['A']: # collect building names
		building.append(cell.value)
	
	for cell in double_ws['B']:
		amount.append(cell.value)
	
	for cell in double_ws['C']:
		descrip.append(cell.value)
	
	data = list(zip(amount, descrip))
	data = dict(zip(building, data))
	return data


# go to the last opened window (MOVE TO QB)
def get_qb():
	keyDown('alt')
	press('tab')
	press('tab')
	keyUp('alt')

# START ENTERING INTO QB
# print the info into the program focused and press down and do again
def print_single():
	for file in single_list:
		print(file)
		time.sleep(.02)
		if type(file)==str:
			pyperclip.copy(file)
			time.sleep(1)
			hotkey('ctrl', 'v')
		else:
			typewrite(str(file))
		if file == 0:
			press('tab')
			typewrite(str(file))
			hotkey('shift','tab')
		press('down')

def print_double(): # for bills in QB
	for building in double_dict:
		amount = double_dict[building][0]
		descrip = double_dict[building][1]
		pyperclip.copy(building)
		hotkey('ctrl','v') # paste building
		press('tab')
		typewrite(str(amount))
		press('tab')
		time.sleep(.3)
		typewrite(descrip)
		press('tab')
		press('tab')
		press('tab')
		time.sleep(.3)

def print_gl_detail():
	for item in gl_dict:
		facility = gl_dict[item][0] # fac name - MUST BE AS INTERCOMPANY
		description = gl_dict[item][1] # description
		debit = gl_dict[item][2]
		credit = gl_dict[item][3]
		print(facility)
		print(debit)
		print(credit)
		pyperclip.copy(facility)
		hotkey('ctrl','v')
		if debit != 0: 
			press('tab')
			press('tab')
			typewrite(str(debit))
			time.sleep(.5)
		else: 
			press('tab')
			typewrite(str(credit))
			press('tab')
			# press('tab')
			time.sleep(.5)
		typewrite(str(description))
		press('tab')
		press('tab')
		press('tab')
		time.sleep(.5)
		
gl_dict = get_gl_info() #use for entering GL detail
# single_list = get_single_info()

get_qb() # focus QB
print_gl_detail() 
# print_single()
