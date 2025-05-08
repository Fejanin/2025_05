import openpyxl

workbook = openpyxl.load_workbook("test.xlsx")

sheet = workbook.active  # Получение активного листа
sheet = workbook["Лист1"]  # Получение листа по имени

rows = sheet.max_row
columns = sheet.max_column

for row in range(1, rows + 1):
    for cell in sheet[row]:
        print((cell.value, cell), end=' ')
    print('\n')

print(dir(cell))
print(f'''{cell.alignment = }||\n{cell.base_date = }||\n{cell.border = }||\n{cell.check_error = }||
{cell.check_string = }||\n{cell.col_idx = }||\n{cell.column = }||\n{cell.column_letter = }||
{cell.comment = }||\n{cell.coordinate = }||\n{cell.data_type = }||\n{cell.encoding = }||\n
{cell.fill = }||\n{cell.font = }||\n{cell.has_style = }||\n{cell.hyperlink = }||\n
{cell.internal_value = }||\n{cell.is_date = }||\n{cell.number_format = }||\n{cell.offset = }||\n
{cell.parent = }||\n{cell.pivotButton = }||\n{cell.protection = }||\n{cell.quotePrefix = }||\n
{cell.row = }||\n{cell.style = }||\n{cell.style_id = }||\n{cell.value = }||\n
''')

print('\n', 100 * '#', '\n')

from openpyxl.styles.numbers import BUILTIN_FORMATS


for key, val in BUILTIN_FORMATS.items():
    print(f'{key = }: {val = }')

