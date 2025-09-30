from openpyxl.reader.excel import load_workbook

filename = 'delnone.xlsx'
wb = load_workbook(filename)
# ws = wb['导入幼儿']

wb.save(filename)
wb.close()
