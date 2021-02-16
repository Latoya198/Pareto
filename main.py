from  openpyxl import load_workbook

wb = load_workbook(filename="D:\СППР\Paretto\Банки.xlsx")
sheet = wb.active

for i in range(2, 11):
    if sheet['B'+str(i)].value >= 7 and sheet['C'+str(i)].value >= 7 and sheet['E'+str(i)].value <= 2:
        print(sheet['A'+str(i)].value)