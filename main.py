# Classes #


class Volunteer:
    def __init__(self, name, email, phone, contactMethod, allergies):
        self.name = name
        self.email = email
        self.phone = phone
        self.contactMethod = contactMethod
        self.allergies = allergies


class Group:
    def __init__(self, member1, member2, member3, member4, day1, day2):
        self.members = [member1, member2, member3, member4]
        self.day1 = day1
        self.day2 = day2
        self.allergies = []
        for member in self.members:
            for allergy in member.allergies:
                if allergy not in self.allergies:
                    self.allergies.append(allergy)

    def __len__(self):
        return len(self.members)


# File #

vFile = open("text.txt")
data = []

for line in vFile:
    data.append(line.strip('\n').split('\t'))

data.remove(data[0])

print(data)





