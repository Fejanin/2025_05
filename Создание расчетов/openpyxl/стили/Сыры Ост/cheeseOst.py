from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side
import re


HEADER = 3  # номер заголовочной строки
FIRST_ROW_DATA = 6  # первая строка с данными
MARK2 = 10  # номер колонки "метка2"
HAVE_FORMULS = ('разн', 'вес', 'ср нов', 'кон ост', 'факт')
CELL_BORDER = ('расчет', 'заказ филиала')

step_row = 5  # пропускаем кол-во строк перед формированием новой таблицы

thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))


def get_columns_abc_by_name(name_cols, num_header_row, sheet):
    name_cols_dict = {key: None for key in name_cols}
    for cell in sheet[num_header_row]:
        if cell.value:
            if cell.value in name_cols_dict:
                name_cols_dict[cell.value] = cell.column_letter
    return name_cols_dict


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


# file_name = "new_дв 30,06,25 бррсч ост сыр.xlsx"
file_name = input("Введите название файла: ")
workbook = load_workbook(filename=file_name)
sheet = workbook.active

max_col = get_max_col(row=HEADER)
if not max_col:
    raise Exception("В заголовке содержаться только пустые значения!")
max_row = get_max_row()
step = max_row + step_row
if not max_row:
    raise Exception("В первой колонке только пустые значения!")

original = {}
duplicate = {}

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

sheet.delete_rows(FIRST_ROW_DATA, max_row - FIRST_ROW_DATA + step_row)

abc_cols_with_formuls = get_columns_abc_by_name(HAVE_FORMULS, HEADER, sheet)
# исправляем формулы, согласно строк, в к-ых они находятся
for key in HAVE_FORMULS:
    row = sheet[f'{abc_cols_with_formuls[key]}{FIRST_ROW_DATA}:{abc_cols_with_formuls[key]}{max_row}']
    for cell in row:  # cell -> tuple
        if type(cell[0].value) == str and '=' in cell[0].value:
            cell[0].value = re.sub(r'\d+', str(cell[0].row), cell[0].value)

abc_cols_with_border = get_columns_abc_by_name(CELL_BORDER, HEADER, sheet)
for key in CELL_BORDER:
    row = sheet[f'{abc_cols_with_border[key]}{FIRST_ROW_DATA}:{abc_cols_with_border[key]}{max_row}']
    for cell in row:  # cell -> tuple
        cell[0].border = thin_border

workbook.save(filename=f"res_{file_name}")
