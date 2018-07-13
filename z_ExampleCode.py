import xlwt
import datetime
import sys
import math
import settings
import time
import Variables

#INSTRUCTIONS: DO NOT PUT THE SAME DOOR IN TWICE, JUST BUMP THE QTY

style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
style2 = xlwt.easyxf('font: name Arial,  bold on; align: wrap on, vert centre, horiz center',
    num_format_str='#,##0.00')
style3 = xlwt.easyxf('font: name Calibri, bold on, height 220',
    num_format_str='#,##0.00')
style4 = xlwt.easyxf('font: name Calibri, bold on, height 220; align: wrap on, horiz center',
    num_format_str='#,##0.00')
style5 = xlwt.easyxf('font: name Calibri, height 220; borders: left thin, right thin, top thin, bottom thin; align: wrap on, vert centre, horiz center',
    num_format_str='#,##0')
style6 = xlwt.easyxf('font: name Calibri, height 220; borders: left thin, right thick, top thin, bottom thin; align: wrap on, vert centre, horiz center',
    num_format_str='#,##0')
style7 = xlwt.easyxf('pattern: pattern solid, fore_colour black;')
style8 = xlwt.easyxf('font: name Calibri, bold on, height 220; align: horiz center; borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_colour yellow;',
    num_format_str='#,##0.00')
style9 = xlwt.easyxf('font: name Calibri, bold on, height  400; align: wrap on, horiz center',
    num_format_str='#,##0.00')
style10 = xlwt.easyxf('font: name Calibri, height 220; borders: left thin, right thin, top thin, bottom thin; align: vert centre, horiz center; pattern: pattern solid, fore_colour gray25;',
    num_format_str='#,##0')
style11 = xlwt.easyxf('font: name Calibri, height 220; borders: left thin, right thin, top thin, bottom thin; align: wrap on, vert centre, horiz center; pattern: pattern solid, fore_colour bright_green;',
    num_format_str='#,##0')
style12 = xlwt.easyxf('font: name Calibri, height 220; borders: left thin, right thin, top thin, bottom thin; align: wrap on, vert centre, horiz center',
    num_format_str='@') # decimal places



wb = xlwt.Workbook()
ws = wb.add_sheet('GP 100 DATA')

ws.col(1).width = int(13.5*260)
ws.col(2).width = int(13.5*260)
ws.col(3).width = int(12*260)
ws.col(4).width = int(23*260)
ws.col(5).width = int(12*260)
ws.col(6).width = int(22.5*260)
ws.col(7).width = int(12*260)
ws.col(8).width = int(23*260)

ws.col(11).width = int(13.5*260)
ws.col(12).width = int(17*260)
ws.col(13).width = int(25.5*260)
ws.col(14).width = int(14*260)

ws.row(1).height_mismatch = True
ws.row(1).height = 2*260

#Data Start

#ws.write(1, 0, "hello", style2)

ws.write(1,1, 'GP 100 DATA', style9)
ws.merge(1, 1, 1, 8)

ws.write(3,1, 'DATA', style4)
ws.merge(3, 3, 1, 2)
#Top cell, bottom cell, left cell, right cell)

ws.write(3,4, 'SHEETS OF MATERIAL', style4)

ws.write(3,6, 'TIME TO RUN FULL QTY', style4)

ws.write(3,8, 'HARDWARE', style4)

ws.write(3,11, 'PYTHON DATA', style4)
ws.merge(3, 3, 11, 14)

ws.write(5,1, "GP 100 Doors", style10)
ws.write(5,2, "Quantity", style10)
ws.write(5,4, "14GA CR (60x120)", style10)
#ws.write(4,4, "18GA (48x120)", style6)
ws.write(5,6, "Laser Time (min)", style10)
#ws.write(4,6, "Punch Time (min)", style5)
ws.write(5,8, "Cams", style10)

ws.write(5,11, "GP 100 Door", style10)
ws.write(5,12, "Doors Per Sheet", style10)
ws.write(5,13, "Laser Time Per Sheet (min)", style10)
ws.write(5,14, "Cams Per Door", style10)


#End Data


def GP_laser_time():
    x = 6
    while True:
        GP_user_door_input = (input("\n> GP Door: "))

        if GP_user_door_input =="help":
            print("\n Rules:\n")
            print("1) Do not have Stock Calculator Excel Sheet open when using exe")
            print("2) Don't write letters in QTY or exe will crash")
            print("3) Inputs are case-sensitive")
            print("4) Don't use the same door twice")

            print("\n Guidelines:\n")
            print("1) All laser times have +1 minute -- to account for pallet changing")
            print("2) All punch times that only use chute have + 3 minutes -- to account for preparation")
            print("3) All punch times that use the chute and tabs have + 4 minutes -- to account for preparation and partial tab cuts")
            print("4) All punch times that use tabs only have + 5 minutes -- to account for preparation and full tab cuts\n")

            print("Calculator is sorted by most common (GP, UAD, Basic), and then alphabetical (AL, Classic, etc.)")


            GP_laser_time()

        if GP_user_door_input == "done":

#            print("\nYour GP door list is", settings.gp_doorlist[:])
#            print("Your Qty List is", settings.gp_qty_list[:])

            user_input_dictionary = dict(zip(settings.gp_doorlist, settings.gp_qty_list))
            print("\nYour Final GP Door Input is", user_input_dictionary)

#            print("Your GP Doors in a sheet is", my_doorspersheet_list[:])
#            print("Your Time to run ONE sheet is", my_runtime_list[:])

#            print("\nEvaluation...\n")

            GP_sheets_ran = [c / t for c,t in zip(settings.gp_qty_list, my_doorspersheet_list)] # dividing lists
            settings.GP_sheets_ran_clean = [int(math.ceil(n)) for n in GP_sheets_ran] # rounding divided sheets
            GP_total_time_per_doororder = [c*t for c,t in zip(settings.GP_sheets_ran_clean, my_runtime_list)]

#            print("Your exact sheets to fill exact qty are", GP_sheets_ran)
#            print("Your actual sheets ran are", settings.GP_sheets_ran_clean)
#            print("Your run time list in minutes is", GP_total_time_per_doororder)
            settings.GP_total_runtime = sum(GP_total_time_per_doororder[:])
#            print("Your total laser runtime is", settings.GP_total_runtime, "minutes")

            GP_total_cams = [c*t for c,t in zip(my_cams_list, settings.gp_qty_list)]
            settings.GP_cams = sum(GP_total_cams[:])

#            print(settings.GP_cams, "hello this is GP cams")

            global gpdoors
            gpdoors = settings.gp_doorlist[:]

            global gpqty
            gpqty = settings.gp_qty_list

            global excel_integer
            excel_integer = x

            GP_punch_time()


        if GP_user_door_input == "final":
            wb.save('Stock Calculator.xls')
            import Final_Calculation

        if not GP_user_door_input in gp_runtime_dct:
            print("This Door Doesn't Exist")
            print("Reasons: (1) You made a Typo, (2) you didn't use lowercase 'x' to define your door")
            print("Try Again")



        elif GP_user_door_input in gp_runtime_dct :
            runtime = my_runtime_list.append(gp_runtime_dct[GP_user_door_input]) # Adds doors (24x24) into list)
            doorset = settings.gp_doorlist.append(GP_user_door_input) # A way to describe stuff
            user_qty_input = eval(input("> Qty: "))
            qty = settings.gp_qty_list.append(user_qty_input)
            doorspersheet = my_doorspersheet_list.append(gp_sheet_dct[GP_user_door_input]) # lists doors per sheet for specific door
            cams = my_cams_list.append(gp_cams_dct[GP_user_door_input])

            ws.write(x, 1, GP_user_door_input, style11)
            ws.write(x, 2, user_qty_input, style11)

            excel_doorspersheet = gp_sheet_dct[GP_user_door_input]
            excel_14ga_sheets = user_qty_input / excel_doorspersheet
            excel_14ga_sheets_clean = math.ceil(excel_14ga_sheets)

            ws.write(x, 4, excel_14ga_sheets_clean, style5)

            excel_lasertimepersheet = gp_runtime_dct[GP_user_door_input]
            excel_lasertime_sheets = excel_lasertimepersheet * excel_14ga_sheets_clean

            ws.write(x, 6, excel_lasertime_sheets, style12)

            excel_cams = gp_cams_dct[GP_user_door_input] * user_qty_input

            ws.write(x, 8, excel_cams, style5)

            ws.write(x, 11, GP_user_door_input, style5)
            ws.write(x, 12, gp_sheet_dct[GP_user_door_input], style5)
            ws.write(x, 13, gp_runtime_dct[GP_user_door_input], style12)
            ws.write(x, 14, gp_cams_dct[GP_user_door_input], style5)

            x = x + 1






    GP_laser_time()

def GP_punch_time ():

    x = excel_integer + 1

    gp_punch_array = ["6x6", "8x8", "10x10", "12x12", "14x14", "16x16", "18x18", "20x20", "22x22", "24x24", "26x26", "28x28", "30x30", "32x32", "34x34", "36x36", "38x38", "40x40", "42x42", "44x44", "46x46", "48x48"]

#    print("\nI'm in the punch_time function", gpdoors) # because I made a global variable, I can now use "gpdoors" here
#    print("my quantity list", gpqty)

    dictionary = dict(zip(gpdoors, gpqty))
#    print(dictionary) # This makes a dictionary for ['Type of Door' : 'Qty']

    ws.write(x, 1, "GP 100 RFs", style10)
    ws.write(x, 2, "Quantity", style10)
    ws.write(x, 4, "18GA CR (48x120)", style10)
    ws.write(x, 6, "Punch Time (min)", style10)

    ws.write(x, 11, "GP 100 RFs", style10)
    ws.write(x, 12, "RFs Per Sheet", style10)
    ws.write(x, 13, "Punch Time Per Sheet (min)", style10)

    x = x + 1



    if '6x6' in gpdoors:
        gpqty_rf = dictionary['6x6'] * 4
#        print("Total qty of 6x6", gpqty_rf)
        my_counter[0] = my_counter[0] + gpqty_rf

    if '8x8' in gpdoors:
        gpqty_rf = dictionary['8x8'] * 4
#        print("Total qty of 8x8", gpqty_rf)
        my_counter[1] = my_counter[1] + gpqty_rf

    if '10x10' in gpdoors:
        gpqty_rf = dictionary['10x10'] * 4
#        print("Total qty of 8x8", gpqty_rf)
        my_counter[2] = my_counter[2] + gpqty_rf

    if '12x12' in gpdoors:
        gpqty_rf = dictionary['12x12'] * 4
#        print("Total qty of 12x12", gpqty_rf)
        my_counter[3] = my_counter[3] + gpqty_rf

    if '14x14' in gpdoors:
        gpqty_rf = dictionary['14x14'] * 4
#        print("Total qty of 14x14", gpqty_rf)
        my_counter[4] = my_counter[4] + gpqty_rf

    if '16x16' in gpdoors:
        gpqty_rf = dictionary['16x16'] * 4
#        print("Total qty of 16x16", gpqty_rf)
        my_counter[5] = my_counter[5] + gpqty_rf

    if '18x18' in gpdoors:
        gpqty_rf = dictionary['18x18'] * 4
#        print("Total qty of 18x18", gpqty_rf)
        my_counter[6] = my_counter[6] + gpqty_rf

    if '20x20' in gpdoors:
        gpqty_rf = dictionary['20x20'] * 4
#        print("Total qty of 20x20", gpqty_rf)
        my_counter[7] = my_counter[7] + gpqty_rf

    if '22x22' in gpdoors:
        gpqty_rf = dictionary['22x22'] * 4
#        print("Total qty of 22x22", gpqty_rf)
        my_counter[8] = my_counter[8] + gpqty_rf

    if '24x24' in gpdoors:
        gpqty_rf = dictionary['24x24'] * 4
#        print("Total qty of 24x24", gpqty_rf)
        my_counter[9] = my_counter[9] + gpqty_rf

    if '26x26' in gpdoors:
        gpqty_rf = dictionary['26x26'] * 4
#        print("Total qty of 26x26", gpqty_rf)
        my_counter[10] = my_counter[10] + gpqty_rf

    if '28x28' in gpdoors:
        gpqty_rf = dictionary['28x28'] * 4
#        print("Total qty of 28x28", gpqty_rf)
        my_counter[11] = my_counter[11] + gpqty_rf

    if '30x30' in gpdoors:
        gpqty_rf = dictionary['30x30'] * 4
#        print("Total qty of 30x30", gpqty_rf)
        my_counter[12] = my_counter[12] + gpqty_rf

    if '32x32' in gpdoors:
        gpqty_rf = dictionary['32x32'] * 4
#        print("Total qty of 32x32", gpqty_rf)
        my_counter[13] = my_counter[13] + gpqty_rf

    if '34x34' in gpdoors:
        gpqty_rf = dictionary['34x34'] * 4
#        print("Total qty of 34x34", gpqty_rf)
        my_counter[14] = my_counter[14] + gpqty_rf

    if '36x36' in gpdoors:
        gpqty_rf = dictionary['36x36'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[15] = my_counter[15] + gpqty_rf

    if '38x38' in gpdoors:
        e = 16
        gpqty_rf = dictionary['38x38'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf

    if '40x40' in gpdoors:
        e = 17
        gpqty_rf = dictionary['40x40'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf

    if '42x42' in gpdoors:
        e = 18
        gpqty_rf = dictionary['42x42'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf

    if '44x44' in gpdoors:
        e = 19
        gpqty_rf = dictionary['44x44'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf

    if '46x46' in gpdoors:
        e = 20
        gpqty_rf = dictionary['46x46'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf

    if '48x48' in gpdoors:
        e = 21
        gpqty_rf = dictionary['48x48'] * 4
#        print("Total qty of 36x36", gpqty_rf)
        my_counter[e] = my_counter[e] + gpqty_rf





    # 6x6 0, 8x8 1, 10x10 2, 12x12 3, 14x14 4, 16x16 5, 18x18 6, 20x20 7, 22x22 8, 24x24 9, 26x26 10, 28x28 11, 30x30 12, 32x32 13, 34x34 14, 36x36 15 ---- PLACEMENT IN ARRAY

    if '24x12' in gpdoors:
        gpqty_rf = dictionary['24x12'] * 2
        my_counter[9] = my_counter[9] + gpqty_rf
        my_counter[3] = my_counter[3] + gpqty_rf

    if '24x36' in gpdoors:
        gpqty_rf = dictionary['24x36'] * 2
        my_counter[9] = my_counter[9] + gpqty_rf
        my_counter[15] = my_counter[15] + gpqty_rf



#    print(my_counter)
#    print(gp_punch_qty)

    GP_sheets_ran_punch = [c / t for c, t in zip(my_counter, gp_punch_qty)]  # dividing lists
#    print("Use this array as your exact number of sheets", GP_sheets_ran_punch)

    GP_sheets_ran_clean_punch = [int(math.ceil(n)) for n in GP_sheets_ran_punch]
#    print("This is the clean number of sheets", GP_sheets_ran_clean_punch)

    settings.GP_total_sheets_ran_punch = sum(GP_sheets_ran_clean_punch)
#    print("\n Your total sheets to be punched is", settings.GP_total_sheets_ran_punch, "sheets")

    GP_sheet_time_punch = [c * t for c, t in zip(GP_sheets_ran_clean_punch, gp_punch_qty_time)]  # multiplying lists
#    print("Use this array as your time!", GP_sheet_time_punch)

    settings.GP_total_runtime_punch = sum(GP_sheet_time_punch[:])
#    print("\n Your total punch time is", settings.GP_total_runtime_punch, "minutes")

    #ws.write(x, 5, excel_lasertime_sheets, style5)

#    print("\n HERE IS THE INFO I WNANT IN MY EXCEL SHEET\n")

#    print(gpdoors)
    #print(gp_punch_array)

    dictionary_qty_rfs = dict(zip(gp_punch_array, my_counter))
    dictionary_sheets_rfs = dict(zip(gp_punch_array, GP_sheets_ran_clean_punch))
    dictionary_time_rfs = dict(zip(gp_punch_array, GP_sheet_time_punch))


#    print(dictionary_qty_rfs)
#    print(dictionary_sheets_rfs)
#    print(dictionary_time_rfs)

    if dictionary_qty_rfs['6x6'] != 0:
        y = '6x6'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[0], style5)
        ws.write(x, 13, gp_punch_qty_time[0], style12)
        x = x + 1

    if dictionary_qty_rfs['8x8'] != 0:
        y = '8x8'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[1], style5)
        ws.write(x, 13, gp_punch_qty_time[1], style12)
        x = x + 1

    if dictionary_qty_rfs['10x10'] != 0:
        y = '10x10'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[2], style5)
        ws.write(x, 13, gp_punch_qty_time[2], style12)
        x = x + 1

    if dictionary_qty_rfs['12x12'] != 0:
        y = '12x12'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[3], style5)
        ws.write(x, 13, gp_punch_qty_time[3], style12)
        x = x + 1

    if dictionary_qty_rfs['14x14'] != 0:
        y = '14x14'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[4], style5)
        ws.write(x, 13, gp_punch_qty_time[4], style12)
        x = x + 1

    if dictionary_qty_rfs['16x16'] != 0:
        y = '16x16'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[5], style5)
        ws.write(x, 13, gp_punch_qty_time[5], style12)
        x = x + 1

    if dictionary_qty_rfs['18x18'] != 0:
        y = '18x18'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[6], style5)
        ws.write(x, 13, gp_punch_qty_time[6], style12)
        x = x + 1

    if dictionary_qty_rfs['20x20'] != 0:
        y = '20x20'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[7], style5)
        ws.write(x, 13, gp_punch_qty_time[7], style12)
        x = x + 1

    if dictionary_qty_rfs['22x22'] != 0:
        y = '22x22'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[8], style5)
        ws.write(x, 13, gp_punch_qty_time[8], style12)
        x = x + 1

    if dictionary_qty_rfs['24x24'] != 0:
        y = '24x24'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[9], style5)
        ws.write(x, 13, gp_punch_qty_time[9], style12)

        x = x + 1

    if dictionary_qty_rfs['26x26'] != 0:
        y = '26x26'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[10], style5)
        ws.write(x, 13, gp_punch_qty_time[10], style12)
        x = x + 1

    if dictionary_qty_rfs['28x28'] != 0:
        y = '28x28'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[11], style5)
        ws.write(x, 13, gp_punch_qty_time[11], style12)
        x = x + 1

    if dictionary_qty_rfs['30x30'] != 0:
        y = '30x30'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[12], style5)
        ws.write(x, 13, gp_punch_qty_time[12], style12)
        x = x + 1

    if dictionary_qty_rfs['32x32'] != 0:
        y = '32x32'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[13], style5)
        ws.write(x, 13, gp_punch_qty_time[13], style12)
        x = x + 1

    if dictionary_qty_rfs['34x34'] != 0:
        y = '34x34'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[14], style5)
        ws.write(x, 13, gp_punch_qty_time[14], style12)
        x = x + 1

    if dictionary_qty_rfs['36x36'] != 0:
        y = '36x36'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[15], style5)
        ws.write(x, 13, gp_punch_qty_time[15], style12)
        x = x + 1

    if dictionary_qty_rfs['38x38'] != 0:
        y = '38x38'
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[16], style5)
        ws.write(x, 13, gp_punch_qty_time[16], style12)
        x = x + 1

    if dictionary_qty_rfs['40x40'] != 0:
        y = '40x40'
        z = 17
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[z], style5)
        ws.write(x, 13, gp_punch_qty_time[z], style12)
        x = x + 1

    if dictionary_qty_rfs['42x42'] != 0:
        y = '42x42'
        z = 18
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[z], style5)
        ws.write(x, 13, gp_punch_qty_time[z], style12)
        x = x + 1

    if dictionary_qty_rfs['44x44'] != 0:
        y = '44x44'
        z = 19
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[z], style5)
        ws.write(x, 13, gp_punch_qty_time[z], style12)
        x = x + 1

    if dictionary_qty_rfs['46x46'] != 0:
        y = '46x46'
        z = 20
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[z], style5)
        ws.write(x, 13, gp_punch_qty_time[z], style12)
        x = x + 1

    if dictionary_qty_rfs['48x48'] != 0:
        y = '48x48'
        z = 21
        ws.write(x, 1, y, style5)
        ws.write(x, 2, dictionary_qty_rfs[y], style5)
        ws.write(x, 4, dictionary_sheets_rfs[y], style5)
        ws.write(x, 6, dictionary_time_rfs[y], style12)
        ws.write(x, 11, y, style5)
        ws.write(x, 12, gp_punch_qty[z], style5)
        ws.write(x, 13, gp_punch_qty_time[z], style12)
        x = x + 1





    #Final Result

    x = x + 1

    ws.write(x, 0, '', style7)
    ws.merge(x, x, 0, 15)

    x = x + 2

    ws.write(x, 4, 'Total Material Required For GP Order', style8)
    ws.merge(x, x, 4, 5, style8)

    x = x + 1

    ws.write(x, 4, '14GA CR (60x120)', style5)
    ws.write(x, 5, sum(settings.GP_sheets_ran_clean), style5)

    x = x + 1

    ws.write(x, 4, '18GA CR (48x120)', style5)
    ws.write(x, 5, settings.GP_total_sheets_ran_punch, style5)

    x = x + 2

    ws.write(x, 4, 'Total Laser Time (h:m.s)', style8)
    ws.write(x, 6, 'Laser Time + 10% (h:m.s)', style8)
    ws.write(x, 8, 'Total Cams', style8)

    x = x + 1


    hours = int(settings.GP_total_runtime / 60)
    minutes = (settings.GP_total_runtime) % 60
    seconds = (settings.GP_total_runtime * 60) % 60

    settings.h_mm_GP_total_runtime = ("%d:%02d.%02d" % (hours, minutes, seconds))

    ws.write(x, 4, settings.h_mm_GP_total_runtime, style5)

    hours = int(settings.GP_total_runtime * 1.1 / 60)
    minutes = (settings.GP_total_runtime * 1.1) % 60
    seconds = (settings.GP_total_runtime * 1.1 * 60) % 60

    h_mm_GP_10per_runtime = ("%d:%02d.%02d" % (hours, minutes, seconds))

    ws.write(x, 6, h_mm_GP_10per_runtime, style5)

    ws.write(x, 8, settings.GP_cams, style5)

    x = x + 2

    ws.write(x, 4, 'Total Punch Time (h:m.s)', style8)
    ws.write(x, 6, 'Punch Time + 10% (h:m.s)', style8)

    x = x + 1

    hours = int(settings.GP_total_runtime_punch / 60)
    minutes = (settings.GP_total_runtime_punch) % 60
    seconds = (settings.GP_total_runtime_punch * 60) % 60

    settings.h_mm_GP_total_runtime_punch = ("%d:%02d.%02d" % (hours, minutes, seconds))

    ws.write(x, 4, settings.h_mm_GP_total_runtime_punch, style5)

    hours = int(settings.GP_total_runtime_punch * 1.1 / 60)
    minutes = (settings.GP_total_runtime_punch) *1.1 % 60
    seconds = (settings.GP_total_runtime_punch * 1.1 * 60) % 60

    h_mm_GP_10per_runtime_punch = ("%d:%02d.%02d" % (hours, minutes, seconds))

    ws.write(x, 6, h_mm_GP_10per_runtime_punch, style5)

    wb.save('Stock Calculator.xls')

    import UAD_calculation

#    sys.exit(0)


print("\n---Stock Calculator---\n")

print("Michael McCabe: Aug 2017\n")



my_runtime_list = []
settings.gp_qty_list = []
my_doorspersheet_list = []
my_cams_list = []

settings.gp_doorlist = [] # Just here to describe stuff

# gp_runtime_dct = {'24x24': 12, '22x22': 10, '18x18': 6, '24x12': 8} #Sheet of doors: minutes
# gp_sheet_dct = {'24x24': 8, '22x22': 12, '18x18': 15, '24x12': 10} #Sheet of doors: amount per sheet
# gp_cams_dct = {'24x24': 4, '22x22': 3, '18x18': 3, '24x12': 3} #Sheet of doors: amount per sheet

gp_runtime_dct = {'6x6': 20.4, '8x8': 17.2, '10x10': 18.8, '12x12': 14.9, '14x14': 14.3, '16x16': 14.6, '18x18': 14.1, '20x20': 10.7, '22x22': 9.9, '24x24': 10.4,  '28x28': 5.2, '30x30': 5.5, '36x36': 4.5, '48x48': 5.7, '24x36' : 6.6              } #Sheet of doors: minutes
gp_sheet_dct = {'6x6': 55, '8x8': 36, '10x10': 32, '12x12': 21, '14x14': 18, '16x16': 15, '18x18': 10, '20x20': 10, '22x22': 8, '24x24': 8, '28x28': 3, '30x30': 3, '36x36': 2, '48x48': 2,                             '24x36' : 4                } #Sheet of doors: amount per sheet
gp_cams_dct = {'6x6': 1, '8x8': 1, '10x10': 3, '12x12': 3, '14x14': 3, '16x16': 3, '18x18': 3, '20x20': 3, '22x22': 4, '24x24': 4, '28x28': 6, '30x30': 6, '36x36': 7, '48x48': 12,                                     '24x36' : 5                } #Cams Per Door






#### Defining entire variables

#### END Defining entire variables



# 6x6 0, 8x8 1, 10x10 2, 12x12 3, 14x14 4, 16x16 5, 18x18 6, 20x20 7, 22x22 8, 24x24 9, 26x26 10, 28x28 11, 30x30 12, 32x32 13, 34x34 14, 36x36 15 ---- PLACEMENT IN ARRAY
# 6x6 130, 8x8 120, 10x10 110, 12x12 100, 14x14 90, 16x16 85, 18x18 80, 20x20 75, 22x22 70, 24x24 65, 26x26 60, 28x28 55, 30x30 50, 32x32 45, 34x34 40, 36x36 35 ----QTY IN A SHEET
# 6x6 18 8x8 17, 10x10 16.5, 12x12 16, 14x14 15.5, 16x16 15, 18x18 14.5, 20x20 14, 22x22 13.5, 24x24 13, 26x26 12.5, 28x28 12, 30x30 11.5, 32x32 11, 34x34 10.5, 36x36 10 ----TIME PER SHEET

gp_punch_qty = (000.1, 132, 99, 96, 66, 66, 66, 55, 54, 51, 00.1, 43, 38, 00.1, 00.1, 36, 00.1, 00.1, 00.1, 00.1, 00.1, 22)
gp_punch_qty_time = (000.1, 29.2, 23.6, 23.9, 19.3, 20.2, 21.1, 14.3, 15.7, 16.3, 00.1, 15.1, 14.5, 00.1, 00.1, 14.9, 00.1, 00.1, 00.1, 00.1, 00.1, 13.9)

my_counter = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

print("Add a GP door and QTY, write 'help' for more information, write 'done' when finished")
print("GP Entry List: 8x8 10x10 12x12 14x14 16x16 18x18 20x20 22x22 24x24 28x28 30x30 36x36 48x48 | 24x36")

GP_laser_time()




