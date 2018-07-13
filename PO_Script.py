# Notes
# ws.row(1).height_mismatch = True - use if row height doesn't match font height
# ALWAYS RUN ON PO_Script.py ... NOT PO_Styles !!

import xlwt
import xlrd
import datetime, xlrd
import math

# Reading/copying values

workbook = xlrd.open_workbook('input.xlsx')
sheet = workbook.sheet_by_index(0)

read_MPO_Number = sheet.cell_value(2,1)
read_MPO_Number = int(read_MPO_Number) # Takes out the decimal place of PO#
read_WB_or_SF = sheet.cell_value(2,2)
read_Date = sheet.cell_value(2,3) # This is just a number

# Converting the read_Date number to the full date

read_Date_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(read_Date, workbook.datemode))

# Converting the read_Date full date to a string
read_Date_as_datetime_string = str(read_Date_as_datetime)

# Deleting bits of string
Date_Month = read_Date_as_datetime_string.replace("00:00:00","")
Date_Month = Date_Month.replace("2018-0","") # Months 1-9

#Author Note, make a string replacement here if you want to get rid of "0#" as the day for the output

Date_Month = Date_Month.replace("2018-","") # Months 10-12

# Here is print(Date_Month, "Date_Month")
Date_Month = Date_Month.replace("-","/") # slash


# Reading Data into Arrays

# The full column information, until it hits nothing
read_PO_data = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]
read_product_data = [sheet.cell_value(row, 2) for row in range(sheet.nrows)]
read_qty_data = [sheet.cell_value(row, 3) for row in range(sheet.nrows)]
read_price_data = [sheet.cell_value(row, 4) for row in range(sheet.nrows)]
read_gauge_data = [sheet.cell_value(row, 7) for row in range(sheet.nrows)]
read_material_data = [sheet.cell_value(row, 8) for row in range(sheet.nrows)]
read_size_data = [sheet.cell_value(row, 9) for row in range(sheet.nrows)]
read_spec_product_data = [sheet.cell_value(row, 11) for row in range(sheet.nrows)]
read_spec_height_data = [sheet.cell_value(row, 12) for row in range(sheet.nrows)]

# Deleting unecessary infromation in array for all data
read_PO_data = read_PO_data[5:]
read_product_data = read_product_data[5:]
read_qty_data = read_qty_data[5:]
read_price_data = read_price_data[5:]
read_gauge_data = read_gauge_data[5:]
read_material_data = read_material_data[5:]
read_size_data = read_size_data[5:]
read_spec_product_data = read_spec_product_data[5:]
read_spec_height_data = read_spec_height_data[5:]

# Number of Products, which yields how long the range will be to iterate
data_range = len(read_PO_data)


# PO Price Calculation - QTY and Price

# Multiplying both arrays to create one array   
data_times_price_array = [read_qty_data*read_price_data for read_qty_data,read_price_data in zip(read_qty_data,read_price_data)]

# Summing the combined array, the total price
read_total_price = sum(data_times_price_array)

# Taking out the decimal for the above value
# The Total Price (Exact)
read_total_price = int(read_total_price)

# The Total Price (Use for PO)
read_total_price_ceil = int(math.ceil(read_total_price / 100))*100



# Title of PO, creating the variables into strings

po_header = "MPO#" + str(read_MPO_Number) + " - " + str(read_WB_or_SF) + " - " + str(Date_Month) + "- $" + str(read_total_price_ceil)

# This is just an idea: read_data = [sheet.cell_value(5, col) for col in range(sheet.ncols)]



# WRITING THE PO

wb = xlwt.Workbook()
ws = wb.add_sheet('PO To Floor')

xlwt.add_palette_colour("custom_gray", 0x21)
wb.set_colour_RGB(0x21, 231, 230, 230)

import PO_Styles

ws.col(0).width = 935 #A 
ws.col(1).width = 3100 #B 
ws.col(2).width = 7150 #C 
ws.col(3).width = 3250 #D 
ws.col(4).width = 3220 #E 
ws.col(5).width = 2700 #F 
ws.col(6).width = 2720 #G
ws.col(7).width = 3020 #H


# set height of cells
# **** Thinking about using nrows instead of a number....
for x in range(0,100):
    ws.row(x).height_mismatch = True
    ws.row(x).height = 287



# Creating the Header of the PO

# ws.write(y,x)
ws.write(1, 2, po_header, PO_Styles.style_header)
# ws.merge(start-y,final-y,start-x,final-x)
ws.merge(1, 2, 2, 6, PO_Styles.style_header)


# Product Headers (Don't Change)

ws.write(4, 1, 'P.O.#', PO_Styles.style_gray_fill_left)
ws.write(4, 2, 'Product', PO_Styles.style_gray_fill_left)
ws.write(4, 3, 'QTY Needed', PO_Styles.style_gray_fill_left)
ws.write(4, 4, 'QTY Made', PO_Styles.style_gray_fill_left)
ws.write(4, 5, 'Supervisor Sign Off	', PO_Styles.style_gray_fill_left)
ws.merge(4, 4, 5, 6, PO_Styles.style_gray_fill_left)

# Product Data (Does Change)


row_x = 5 # row_x is the row variable 

# x will be repeated based on data_range (number of products listed)
# for numbers, an int will be used to get rid of deicmal
for x in range(data_range):  
    a = int(read_PO_data[x])
    b = read_product_data[x]
    c = read_qty_data[x]
    ws.write(row_x, 1, a, PO_Styles.style_normal_small_center)
    ws.write(row_x, 2, b, PO_Styles.style_normal_left)
    ws.write(row_x, 3, c, PO_Styles.style_normal_center)
    ws.write(row_x, 4, "", PO_Styles.style_normal_center)
    ws.write(row_x, 5, "", PO_Styles.style_normal_center)
    ws.merge(row_x, row_x, 5, 6, PO_Styles.style_normal_center)
    row_x = row_x + 1
else: # if loop ends, add another +2 to row_x
    row_x = row_x + 2

ws.write(row_x, 1, "MATERIAL QUANTITIES", PO_Styles.style_header_normal)
ws.merge(row_x, row_x, 1, 6, PO_Styles.style_header_normal)



'''

'''




wb.save('PO_Script.xls')