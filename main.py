from openpyxl import load_workbook

wb = load_workbook(filename="D:\СППР\Paretopy_pr_1\Банки.xlsx")
sheet = wb.active

percent = (int(input("Процент+ ")))
rait = (int(input("Райтинг+ ")))
msum = (int(input("Мин Сума - ")))
msrok = (int(input("Мин Срок- ")))

cnt = 2
while cnt != 9:
    for i in range(cnt, 10):
        if sheet['B' + str(i)].value >= sheet['B' + str(i + 1)].value and sheet['C' + str(i)].value >= sheet[
            'C' + str(i + 1)].value \
                and sheet['D' + str(i)].value <= sheet['D' + str(i + 1)].value and \
                sheet['E' + str(i)].value <= sheet['E' + str(i + 1)].value:
            print(sheet['A' + str(i)].value)
        else:
            print("Н")
    cnt += 1
    print("#" * 30)

for i in range(2, 11):
    if sheet['B' + str(i)].value >= rait and sheet['C' + str(i)].value >= percent and sheet['E' + str(i)].value <= msrok \
            and sheet['D' + str(i)].value <= msum:
        print(sheet['A' + str(i)].value)


