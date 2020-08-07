from openpyxl import load_workbook
wb = load_workbook('test.xlsx')

ws = wb.active

while(True):
    number = input("號碼: ")
    grades = input("分數: ")
    if number == 'end' and grades == 'end':
        break
    else:
        ws['A'+ number] = grades
        print('\n')
ws.delete_cols(1,1)
wb.save('test.xlsx')


