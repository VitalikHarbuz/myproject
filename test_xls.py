import xlsxwriter

workbook = xlsxwriter.Workbook('report.xlsx')

worksheet = workbook.add_worksheet('ЗВІТ') 

data = [['aaa','bbb','ccc'], [2,3,5]]

worksheet.write_column('C6', data[0])

workbook.close()