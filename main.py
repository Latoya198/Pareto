from openpyxl import load_workbook

wb = load_workbook(filename="D:\СППР\Paretopy_pr_1\Банки.xlsx")
sheet = wb.active

# Алгоритм МКО по Парето
cnt = 2
while cnt < 10:
    for i in range(cnt, 10):
        if sheet['B' + str(cnt)].value >= sheet['B' + str(i + 1)].value and \
                sheet['C' + str(cnt)].value >= sheet['C' + str(i + 1)].value and \
                sheet['D' + str(cnt)].value <= sheet['D' + str(i + 1)].value and \
                sheet['E' + str(cnt)].value <= sheet['E' + str(i + 1)].value and \
                sheet['F' + str(cnt)].value >= sheet['F' + str(i + 1)].value:
            print(sheet['A' + str(cnt)].value)
        elif sheet['B' + str(cnt)].value <= sheet['B' + str(i + 1)].value and \
                sheet['C' + str(cnt)].value <= sheet['C' + str(i + 1)].value and \
                sheet['D' + str(cnt)].value >= sheet['D' + str(i + 1)].value and \
                sheet['E' + str(cnt)].value >= sheet['E' + str(i + 1)].value and \
                sheet['F' + str(cnt)].value <= sheet['F' + str(i + 1)].value:
            print(sheet['A' + str(i + 1)].value)
        else:
            print("Н")
    print("#" * 30)
    cnt += 1

print(" " * 30)
print("Сужение")
print("#" * 30)

# Сужение
# Лексикографическая оптимизация
print("Лексографическая оптимизация\n")
print("#" * 30)
percent = (int(input("Введите минимальный Процент+ ")))
rait = (int(input("Введите минимальный рейтинг Райтинг+ ")))
msrok = (int(input("Введите максимальный срок влада- ")))
msum = (int(input("Введите минимальную сумму Вклада - ")))

for i in range(2, 10):
    if sheet['C' + str(i)].value >= percent and sheet['B' + str(i)].value >= rait and \
            sheet['E' + str(i)].value <= msrok and sheet['D' + str(i)].value > msum:
        print(sheet['A' + str(i)].value)
