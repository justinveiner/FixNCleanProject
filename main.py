# Classes #
import openpyxl, xlsxwriter

class Volunteer:
    def __init__(self, name, email, phone, contactMethod, allergies):
        self.name = name
        self.email = email
        self.phone = phone
        self.contactMethod = contactMethod
        self.allergies = allergies


class Group:
    def __init__(self, members, day1, day2):
        self.members = members
        self.day1 = day1
        self.day2 = day2
        self.allergies = []
        for member in self.members:
            for allergy in member.allergies:
                if allergy not in self.allergies:
                    self.allergies.append(allergy)

    def __len__(self):
        return len(self.members)
    
    def __add__(self, otherGroup):
        newGroup = Group(self.members + otherGroup.members, self.day1, 4)
        return newGroup

class Client:
    def __init__(self, name, address, email, phone, allergies, timeslot):
        self.name = name
        self.address = address
        self.email = email
        self.phone = phone
        self.allergies = allergies
        self.timeslot = timeslot


# FILE READ #

wb = openpyxl.load_workbook("Copy-of-Volunteer-list-example.xlsx")
sheet = wb["Sheet1"]
cols = ["A","B","C","D"]


#sheet["A1"].va


# CODE #

allGroups = []
for row in range(2,112):
    names = []
    for i in range(1, 15,3):
        j = chr(65 + i)
        name = sheet[j+str(row)].value
        email = sheet[j+str(row+1)].value
        if name != None:
            vol = Volunteer(name,email,'','',[]) #placeholders, no contact info or phone number present
            names.append(vol)
    allGroups.append(Group(names,'',''))



volDict = {}
clientDict = {}


data=[]


for group in allGroups:
    team = []
    for member in group.members:
       team.append(member.name)
    data.append(team)

print(data)
book = xlsxwriter.Workbook("Testfile.xlsx")
sheet = book.add_worksheet()


row = 1
col = 0
headers = ["Team","Saturday Morning","Client","Saturday Afternoon","Client","Sunday Morning","Client","Sunday Afternoon","Client"]




for i in range(len(headers)):
    sheet.write(0,i,headers[i])





book.close()














