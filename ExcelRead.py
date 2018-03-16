import openpyxl

wb = openpyxl.load_workbook("Copy-of-Volunteer-list-example.xlsx")
sheet = wb["Sheet1"]
cols = ["A","B","C","D"]
print(sheet['A1'].value)

col = 65 #example: 65 will be converted to 'A' in ascii to get the contents of each cell
row = 2
