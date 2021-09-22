from oauth2client.service_account import ServiceAccountCredentials
import gspread
from datetime import date
import csv
from tkinter import *

import os
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText

class student:
	def __init__(self,lname,fname,num,hall,room,zipcode,email,dob):
		#Last Name,First Name,Student Number,Hall,Room,Address ZIP,Email,Date of Birth
		self.lastName = lname
		self.firstName = fname
		self.num = num
		self.hall = hall
		self.room = room
		self.floor = room[0]
		self.zipcode = zipcode
		self.email = email
		self.dob = dob

scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json',scope)
client = gspread.authorize(creds)

def loadRoster(fileName = 'Roster.csv'):
	roster = {}
	with open(fileName,'r') as csv_file:
		csv_reader = csv.reader(csv_file)
		for line in csv_reader:
			roster[line[1]+' '+line[0]] = student(line[0],line[1],line[2],line[3],line[4],line[5],line[6],line[7])
	return roster

def sendEmail(studentEmail='sahilpwns@gmail.com'):

	#Need to change OS enviromental variables to laurel log-in / pw
	#Need to allow less secure applications 

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



def AddName(roster = None, name = 'Maria Pepper',deliveryCompany = 'Unknown',si = 'NA',):
	if ',' in name:

		name = name.split(',')
		deliveryCompany = name[1].strip()
		name = name[0]

	roster = loadRoster() if not roster else roster
	if name in roster:
		packageSheet = client.open('Temp Sheet').worksheet(roster[name].floor)
		# i = 1
		# while packageSheet.cell(i,2).value != None: i+=1
		packageSheet.update_cell(1,13,'=MATCH("@",ARRAYFORMULA(B2:B&"@"),0)')
		i = int(packageSheet.cell(1,13).value)+1
		packageSheet.update_cell(1,13,'')
		packageSheet.update_cell(i,1,str(date.today()))
		packagenum = roster[name].floor + '-' + str(i).zfill(3)
		packageSheet.update_cell(i,2,packagenum)
		packageSheet.update_cell(i,3,deliveryCompany)
		packageSheet.update_cell(i,4,name)
		packageSheet.update_cell(i,5,roster[name].room)
		packageSheet.update_cell(i,6,si)
		packageSheet.update_cell(i,7,'Y' if sendEmail() else 'N')
		packageSheet.update_cell(i,8,roster[name].num)
		return f'Package inputted! Package Number for {roster[name].firstName}: {packagenum}'
	else:
		return 'Name not found, please check spelling'


root = Tk()
Name = Entry(root, width = 50)
Name.pack()


def myClick():
	name = Name.get()
	#print(f'Looking for name {name}')
	
	myLabel = Label(root, text=AddName(None,name))
	myLabel.pack()

	#print(f'Coming from company {DeliveryCompany}')
	#myLabel = Label(root, text=e.get())
	#myLabel.pack()

myButton = Button(root, text='Enter Student Name for Package Entry', command=myClick)
myButton.pack()


root.mainloop()


	

# roster = loadRoster()
# for x in ['Amber Eiserle', 'Josef Birman', 'Cindy Yang']:
# 	print(f'Adding {x}...')
# 	AddName(roster,x)


# python_test = client.open('Temp Sheet').sheet1
# python_test = client.open('Temp Sheet').worksheet("Roster")