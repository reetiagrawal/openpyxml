#READ DATA FROM ROW 1 AND COLUMN 1
'''import openpyxl
path = "/home/shiv/Downloads/openxml.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
cell_obj = sheet_obj.cell(row = 1, column= 1)
print(cell_obj.value)'''

#READ MAXIMUM ROWS AND COLUMNS
'''import openpyxl
path = "/home/shiv/Downloads/openxml.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
print(sheet_obj.max_row)
print(sheet_obj.max_column)
'''
#READ ALL COLUMN NAMES
'''import openpyxl
path = "/home/shiv/Downloads/openxml.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
max_col = sheet_obj.max_column

for i in range(1,max_col +1):
    cell_obj = sheet_obj.cell(row = 1,column=i)
    print(cell_obj.value)
'''
'''import openpyxl
path = "/home/shiv/Downloads/openxml.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
m_row = sheet_obj.max_row
for i in range(1,m_row+1):
    cell_obj = sheet_obj.cell(row=i,column=1)
    print(cell_obj.value)
'''
#READ DATA FROM SECOND ROW
'''import openpyxl
path = "/home/shiv/Downloads/openxml.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
max_col = sheet_obj.max_column
for i in range(1,max_col+1):
    cell_obj = sheet_obj.cell(row=2,column=i)
    print(cell_obj.value, end= " ")
'''

