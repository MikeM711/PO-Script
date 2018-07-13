# 

import xlwt # Writes Files
import xlrd # Reads Files

# xlrd documentation - http://xlrd.readthedocs.io/en/latest/api.html

# We will be opening input.xlsx with xlrd to read.  Using "xlrd" in the beginning of the python file allows us to use xlrd of its entirety!
workbook = xlrd.open_workbook('input.xlsx')

# We will be on The first tab of the sheet, it will be called "sheet", since sheet ecompasses workbook, it takes on the xlrd module! As shown below, more specifically, it takes on the workbook, and the sheet page, noted as "sheet_by_index"

sheet = workbook.sheet_by_index(0)

# The data we will be extracting

Test_Data = sheet.cell_value(0, 1)

print("\n", Test_Data, "Test Data")

print("\n", [sheet.cell_value(0, col) for col in range(sheet.ncols)], "idk lol")

# sheet.cell_value = the cell_value METHOD
# Specifically: 
# sheet.cell_value()
# = workbook.sheet_by_index(0)cell_value()
# = xlrd.open_workbook('input.xlsx').sheet_by_index(0).cell_value()
# |xlrd.open_workbook('input.xlsx').sheet_by_index(0)| can be used to replace |sheet|



#sheet.ncols = 8 | it takes the number of variables that are listed in the excel sheet
# col is the range of values that will be using - 0,1,2,3,4,5,6,7
# cell_value is the value of the cell (obviously)
# ^ All of the above are API references (don't know the normal term...)



data = [sheet.cell_value(4,col) for col in range(sheet.ncols)]

# Below is long code that we can use instead - but we really don't want to
# # data = [xlrd.open_workbook('input.xlsx').sheet_by_index(0).cell_value(1, col) for col in range(sheet.ncols)]




# data = [specific_sheet.cell_value(column_position )


workbook = xlwt.Workbook()
sheet = workbook.add_sheet('test')

for index, value in enumerate(data):
    sheet.write(0, index, value)

workbook.save('2-output_test_explanation.xls')