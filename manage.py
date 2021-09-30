from oauth2client.service_account import ServiceAccountCredentials
import gspread
from datetime import date
import csv
from tkinter import *
import os
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
import time

class student:
	def __init__(self,lname,fname,num,hall,room,zipcode,email,dob):
		#Last Name,First Name,Student Number,Hall,Room,Address ZIP,Email,Date of Birth
		self.lastName = lname
		self.firstName = fname
		self.num = num
		self.hall = hall
		self.room = room
		floorNames = {0:'Basement',1:'1st Floor',2:'2nd Floor', 3:'3rd Floor', 4: '4th Floor', 5:'5th Floor',6:'6th Floor',7:'7th Floor'}
		if room!='Room':
			self.floor = floorNames[int(room[0])]
			self.origfloor = room[0] if room[0].isdigit() else '0'
		self.zipcode = zipcode
		self.email = email
		self.dob = dob

def loadRoster(fileName = 'Roster.csv'):
	roster = {}
	with open(fileName,'r') as csv_file:
		csv_reader = csv.reader(csv_file)
		for line in csv_reader:
			roster[line[1]+' '+line[0]] = student(line[0],line[1],line[2],line[3],line[4],line[5],line[6],line[7])
	return roster

def sendEmail(studentEmail='sahilpwns@gmail.com'):

	#Need to change OS enviromental variables to laurel log-in / pw
	#Need to allow less secure applications: https://support.google.com/mail/?p=BadCredentials

	def ReadEmailTemplate(file):
	    oFile = open(file, 'r')
	    Subject = oFile.readline().strip()
	    Body = oFile.read()
	    oFile.close()
	    return [Subject, Body]
	try:
		Subject, Body = ReadEmailTemplate('EmailTemplate.txt')
		msg = EmailMessage()
		BodyM = MIMEText(Body)
		msg.set_content(Body)
		msg['From'] = os.environ['LAUREL_EMAIL']
		msg['To'] = studentEmail
		msg['Subject'] = Subject
		with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
			smtp.login(os.environ['LAUREL_EMAIL'], os.environ['LAUREL_EMAILPW'])
			smtp.sendmail(os.environ['LAUREL_EMAIL'], studentEmail, msg.as_string())

		smtplib.SMTP_SSL('smtp.gmail.com', 465).quit()

		return True
	except:
		return False

def AddName(name,deliveryCompany,initals,roster=None):
	if name=='' or deliveryCompany=='' or initals=='': return 'Please complete all fields'
	roster = loadRoster() if not roster else roster
	if name in roster:
		packageSheet = client.open('Testing Copy (obviously dont use this)').worksheet(roster[name].floor)
		packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(A4:A&"@"),0)+2')
		i = int(packageSheet.cell(1,13).value)
		print(packageSheet.cell(1,13).value.split('-')[0])
		i = int(packageSheet.cell(1,13).value.split('-')[0])+1
		packageSheet.update_cell(1,13,'')
		packageSheet.update_cell(i,1,str(date.today()))
		packagenum = roster[name].origfloor  + '-' + str(i).zfill(3)
		packageSheet.update_cell(i,2,packagenum)
		packageSheet.update_cell(i,3,deliveryCompany)
		packageSheet.update_cell(i,4,name)
		packageSheet.update_cell(i,5,roster[name].room)
		packageSheet.update_cell(i,6,initals)
		#packageSheet.update_cell(i,7,'Y' if sendEmail(roster[name].email) else 'N') 
		packageSheet.update_cell(i,7,'Y' if sendEmail('sahilpwns@gmail.com') else 'N') 
		return f'Package inputted! Package Number for {roster[name].firstName}: {packagenum}'
	else:
		return 'Name not found, please check spelling'

def PackageAdd(roster=None):

	def myClick():
		nonlocal roster
		name = Name.get()
		myLabel = Label(root, text=AddName(roster,name))
		myLabel.pack()

	root = Tk()
	root.title("Sahil's Package Sorting System")
	Name = Entry(root, width = 50)
	Name.pack()
	myButton = Button(root, text='Enter Student Name for Package Entry', command=myClick)
	myButton.pack()
	root.mainloop()

def PickupPackage(name,ID,initals,roster=None):
	if name=='' or ID=='' or initals=='': return 'Please complete all fields'
	roster = loadRoster() if not roster else roster
	if name in roster:
		print(roster[name].floor)
		packageSheet = client.open('Testing Copy (obviously dont use this)').worksheet(roster[name].floor)
		packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(A4:A&"@"),0)+2')
		i = int(packageSheet.cell(1,13).value)
		rateLimiter = 0
		while (i>0):
			if packageSheet.cell(i,4).value==name and packageSheet.cell(i,10).value==None:
				break
			else:
				print(packageSheet.cell(i,10).value)

			i-=1
			rateLimiter+=1
			if rateLimiter>30:
				return 'Package cannot be found, please check manually'
				# rateLimiter = 0
				# print('This package is very old, we are recharging to search again...')
				# time.sleep(50)
				

			
		if i==0: 
			return 'That resident has no package'


		packageSheet.update_cell(i,8,ID)
		packageSheet.update_cell(i,9,str(date.today()))
		packageSheet.update_cell(i,10,initals)
		return f'Package Pickup for {name} has been logged'
	else:
		return 'Name not found, please check spelling'

def getPackageInfo(name,ID,initals,roster=None):
	if name=='' or initals=='': return 'Please complete all fields'
	roster = loadRoster() if not roster else roster
	if name in roster:
		print(roster[name].floor)
		packageSheet = client.open('Testing Copy (obviously dont use this)').worksheet(roster[name].floor)
		packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(A4:A&"@"),0)+2')
		i = int(packageSheet.cell(1,13).value)
		rateLimiter = 0
		while (i>0 or packageSheet.cell(i,10)==None):
			if packageSheet.cell(i,4).value==name:
				break
				#return f'We found {name}'
			i-=1
			rateLimiter+=1
			return 'Package cannot be found, please check manually'
			# if rateLimiter>20:
			# 	rateLimiter = 0
			# 	print('This package is very old, we are recharging to search again...')
			# 	time.sleep(50)
				

			
		if i==0: 
			return 'That resident has no package'
		return f'Package number for {name} is {packageSheet.cell(i,2).value}'

	

scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json',scope)
client = gspread.authorize(creds)


def seperateGUI(roster=None):

	root = Tk()
	win2 = Tk()
	win2.title("Logs")
	root.title("Sahil's Package Sorting System")

	root.geometry("500x300")

	def inputClick():
		nonlocal roster, Name1, DevCompInput, IntialsInput1
		myLabel = Label(win2, text=AddName(Name1.get(),DevCompInput.get(),IntialsInput1.get(),roster)).pack()
		
	def outputClick():
		nonlocal roster, Name2, IDNumInput, IntialsInput2
		myLabel = Label(win2, text=PickupPackage(Name2.get(),IDNumInput.get(),IntialsInput2.get(),roster)).pack()

	def PackageInfoClick():
		nonlocal roster, Name2, IDNumInput, IntialsInput2
		myLabel = Label(win2, text=getPackageInfo(Name2.get(),IDNumInput.get(),IntialsInput2.get(),roster)).pack()

	nameList = [roster[student].firstName + ' ' + roster[student].lastName for student in roster]
	nameList.sort()
	
	nameClicked1 = StringVar()
	nameLabel1 = Label(root, text="Package Input",font=('Helvetica', 18, 'bold')).grid(row=0, column=0)
	nameLabel1 = Label(root, text="Resident Name").grid(row=1, column=0)
	Name1 = Entry(root, textvariable = nameClicked1, width = 20)
	Name1.grid(row=1,column=1)

	NameDropdown1 = OptionMenu(root,nameClicked1,*nameList)
	NameDropdown1.grid(row=1,column=2)

	nameLabel2 = Label(root, text="Delivery Company").grid(row=2, column=0)
	DevCompInput = Entry(root, width = 20)
	DevCompInput.grid(row=2,column=1)

	nameLabel3 = Label(root, text="Intials").grid(row=3, column=0)
	IntialsInput1 = Entry(root, width = 20)
	IntialsInput1.grid(row=3,column=1)

	myButton1 = Button(root, text=' Submit ', command=inputClick).grid(row=4,column=1)

	nameClicked2 = StringVar()
	pickupLabel1 = Label(root, text="Package Pickup",font=('Helvetica', 18, 'bold')).grid(row=5, column=0)
	pickupLabel1 = Label(root, text="Resident Name").grid(row=6, column=0)
	Name2 = Entry(root, textvariable = nameClicked2, width = 20)
	Name2.grid(row=6,column=1)

	NameDropdown2 = OptionMenu(root,nameClicked2,*nameList)
	NameDropdown2.grid(row=6,column=2)

	pickupLabel2 = Label(root, text="Resident ID (on card)").grid(row=7, column=0)
	IDNumInput = Entry(root, width = 20)
	IDNumInput.grid(row=7,column=1)

	pickupLabel3 = Label(root, text="Intials").grid(row=8, column=0)
	IntialsInput2 = Entry(root, width = 20)
	IntialsInput2.grid(row=8,column=1)

	
	myButton3 = Button(root, text=' Get Package Number ', command=PackageInfoClick).grid(row=9,column=1)
	myButton2 = Button(root, text=' Submit ', command=outputClick).grid(row=10,column=1)


	root.mainloop()


seperateGUI(loadRoster())
