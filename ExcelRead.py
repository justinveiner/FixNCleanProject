import openpyxl


wb = openpyxl.load_workbook("Copy-of-Volunteer-list-example.xlsx")
sheet = wb["Sheet1"]
cols = ["A","B","C","D"]
print(sheet['A1'].value)


row = 2

for i in range(65,86):
    row = chr(i)
    names = []
    for j in range(2,15,3):

        name = sheet[row+str(j)].value
        email = sheet[row+str(j+1)].value
        vol = Volunteer(name,email,'','','') #placeholders, no contact info or phone number present
        names.append(vol)
    group = Group(names,sheet[row+'19'].value,sheet[row+'15'].value,sheet[row+'16'].value)
    print(group.names)






