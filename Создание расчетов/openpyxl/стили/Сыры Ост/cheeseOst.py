from openpyxl import load_workbook


HEADER = 3  # номер заголовочной строки
FIRST_ROW_DATA = 6  # первая строка с данными
MARK2 = 10  # номер колонки "метка2"

srip_row = 5


def get_max_col(row):
    max_col = None
    for i in range(1, sheet.max_column):
        if sheet.cell(row=HEADER, column=i).value:
            max_col = i
    return max_col


def get_max_row():
    max_row = None
    for i in range(1, sheet.max_row):
        if sheet.cell(row=i, column=1).value:
            max_row = i
    return max_row


# file_name = "дв 16,06,25 бррсч ост сыр.xlsx"
file_name = input("Введите название файла: ")
workbook = load_workbook(filename=file_name)
sheet = workbook.active

max_col = get_max_col(row=HEADER)
if not max_col:
    raise Exception("В заголовке содержаться только пустые значения!")
max_row = get_max_row()
step = max_row + srip_row
if not max_row:
    raise Exception("В первой колонке только пустые значения!")

original = {}
duplicate = {}
test = {}

for i in range(FIRST_ROW_DATA, max_row + 1):
    if sheet.cell(i, MARK2).value:
        key_name = ' '.join(sheet.cell(i, MARK2).value.lower().split())
        duplicate.setdefault(key_name, []).append(sheet[i][:max_col])
    else:
        key_name = ' '.join(sheet.cell(i, 1).value.lower().split())
        original[key_name] = sheet[i][:max_col]

for i in original:
    some_row = original[i]
    for x in range(1, max_col + 1):
        sheet.cell(step, x).value = some_row[x - 1].value
    step += 1
    if i in duplicate.keys():
        for j in duplicate[i]:
            some_row = j
            for y in range(1, max_col + 1):
                sheet.cell(step, y).value = some_row[y - 1].value
            step += 1

sheet.delete_rows(FIRST_ROW_DATA, max_row - FIRST_ROW_DATA + srip_row)

workbook.save(filename=f"test_{file_name}")
input("It's ok. Push ENTER.")
