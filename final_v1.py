import openpyxl
import os
import xlsxwriter
import smtplib

# Classes #


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
        # automatically collects allergies from the members and assigns to overall group
        for member in self.members:
            for allergy in member.allergies:
                if allergy not in self.allergies:
                    self.allergies.append(allergy)

    def __len__(self):
        return len(self.members)

    def __add__(self, otherGroup):
        newGroup = Group(self.members + otherGroup.members, self.day1, 5)
        return newGroup

    def has_member(self, volunteer):
        return volunteer in self.members


class Client:
    def __init__(self, name, date, address, allergies):
        self.name = name
        self.address = address
        self.allergies = allergies
        self.date = date


# Functions #

def eliminate(volunteer, vDict):
    # This function removes a volunteer from the dictionary
    for key in vDict.keys():
        for group in vDict[key]:
            if group.has_member(volunteer):
                vDict[key].remove(group)
    return vDict


def help():
    # This function prints basic help and tips
    print("----------HELP----------\n"
          "1. Ensure your inputs lack spaces.\n"
          "2. Inputs are case sensitive. Ensure that you are typing inputs with the correct case.\n"
          "3. Ensure that you type the file name correctly.\n"
          "4. Ensure that the volunteer and client information files are in the specified location.\n")


def menu():
    # This code runs the main menu
    print("FIX 'N' CLEAN VOLUNTEER ASSIGNMENT PROGRAM")
    menu_run = True    # Variable Controlling the input check loop

    # This loop runs if the user's input is invalid
    while menu_run == True:

        menu_input = input("----------------------------------------------------------------\n"
                            "Type in the number related to the function you would like to run\n"
                            "1: Volunteer assignment\n2: Automated Email\n3: Help\n4: Quit\n\n->")

        if (menu_input != "1") and (menu_input != "2") and (menu_input != "3") and (menu_input != "4"):
            print("Invalid input. Please type 1, 2, 3, or 4 depending on which function you would like to access.\n")
        else:
            menu_run = False

    return menu_input


def sFormat(string):
    # Repeats asking for a file until a proper one is given
    if len(string) > 5:
        if string[-5:] == ".xlsx":
            if os.path.isfile(string):
                return string
    return sFormat(input("Error: Please re-enter the name of your file, with the file type(i.e. \"example.xlsx\"):"))
    # Need to edit the phrase


def surplus(dict1, dict2, clientDict):
    # Initializes variables storing the shortest list in the array and its index
    # All three dictionaries must have the exact same keys
    shortest = 10000
    # The loop checks for the shortest list in the array and records its index
    for key in dict1.keys():
        # Determines the total amount of members in a section
        tLen = 0
        for group in dict1[key]:
            tLen += len(group)
        for group in dict2[key]:
            tLen += len(group)
        tLen -= len(clientDict[key])
        # Finds the lowest and returns it to be accessed
        if tLen <= shortest:
            shortest = tLen
            best = key

    return best

# VARIABLES #
# Dictionaries

# Places the data into groups
firstChoice = {1: [], 2: [], 3: [], 4: []}
secondChoice = {1: [], 2: [], 3: [], 4: []}
clients = {1: [], 2: [], 3: [], 4: []}
pRun = True # Controls the main program loop


# CODE #

while (pRun == True):
    # Run the Main Menu
    menu_choice = menu()

    # Based on the user's choice, this logical structure determines which functionality is run
    if (menu_choice == 1):
        # Runs the Volunteer Assignment function


    elif (menu_choice == 2):
        # Runs the Email function


    elif (menu_choice == 3):
        # Runs the Help function
        help()

    elif (menu_choice == 4):
        # Ends the progrm loop to close the program
        pRun = False

# Closes the program once the main program loop ends
exit(0)