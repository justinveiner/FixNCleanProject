# Classes #
import openpyxl

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

"""
# File #

#vFile = open("text.txt")
#data = []

#for line in vFile:
    #data.append(line.strip('\n').split('\t'))

#data.remove(data[0])



print(data)

wb = openpyxl.load_workbook("Copy-of-Volunteer-list-example.xlsx")
sheet = wb["Sheet1"]
cols = ["A","B","C","D"]



row = 2

for i in range(65,86):
    row = chr(i)
    names = []
    for j in range(2,15,3):

        name = sheet[row+str(j)].value
        email = sheet[row+str(j+1)].value
        vol = Volunteer(name,email,'','',[]) #placeholders, no contact info or phone number present
        names.append(vol)
    group = Group(names,sheet[row+'15'].value,sheet[row+'16'].value)

    print(group.members[0].name)


"""







