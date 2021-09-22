from oauth2client.service_account import ServiceAccountCredentials
import gspread
from datetime import date
import csv

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

def sendEmail(student):
	return False

def AddName(roster = None, name = 'Maria Pepper',deliveryCompany = 'Unknown',si = 'NA',):
	roster = loadRoster() if not roster else roster
	if name in roster:
		packageSheet = client.open('Temp Sheet').worksheet(roster[name].floor)
		i = 1
		while packageSheet.cell(i,2).value != None: i+=1
		packageSheet.update_cell(i,1,str(date.today()))
		packagenum = roster[name].floor + '-' + str(i).zfill(3)
		packageSheet.update_cell(i,2,packagenum)
		packageSheet.update_cell(i,3,deliveryCompany)
		packageSheet.update_cell(i,4,name)
		packageSheet.update_cell(i,5,roster[name].room)
		packageSheet.update_cell(i,6,si)
		packageSheet.update_cell(i,7,'Y' if sendEmail(roster[name]) else 'N')
		packageSheet.update_cell(i,8,roster[name].num)
	else:
		print('Name not found, please check spelling')
	



# python_test = client.open('Temp Sheet').sheet1
# python_test = client.open('Temp Sheet').worksheet("Roster")