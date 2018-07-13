import xlwt




style_header = xlwt.easyxf('font: name Calibri, height  560; align: wrap on, horiz center, vert center; borders: left medium, right medium, top medium, bottom medium;', num_format_str='#,##0.00')

style_header_normal = xlwt.easyxf('font: name Calibri, bold on, height  220; align: wrap on, horiz center, vert center; borders: left medium, right medium, top medium, bottom medium;', num_format_str='#,##0.00')

style_gray_fill_left = xlwt.easyxf('font: name Calibri, height  220; align: wrap on, horiz left, vert center; pattern: pattern solid, fore_colour custom_gray; borders: left thin, right thin, top thin, bottom thin;', num_format_str='#,##0.00')

style_normal_small_center = xlwt.easyxf('font: name Calibri, height  180; align: wrap on, horiz center, vert center; borders: left thin, right thin, top thin, bottom thin;', num_format_str='###')

style_normal_center = xlwt.easyxf('font: name Calibri, height  220; align: wrap on, horiz center, vert center; borders: left thin, right thin, top thin, bottom thin;', num_format_str='#,###')

style_normal_left = xlwt.easyxf('font: name Calibri, height  220; align: wrap on, horiz left, vert center; borders: left thin, right thin, top thin, bottom thin;', num_format_str='#,##0.00')


# below uses bold
# style_header = xlwt.easyxf('font: name Calibri, bold on, height  600; align: wrap on, horiz center', num_format_str='#,##0.00')