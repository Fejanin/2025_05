import openpyxl
from openpyxl.styles import PatternFill


green_fill = PatternFill(start_color='13E22E', end_color='13E22E', fill_type='solid')


# file_name = 'new_дв 12,12,24 днрсч пок ки.xlsx'
file_name = input('Введите название файла: ')

wb = openpyxl.load_workbook(file_name)

sheet = wb.active  # Получение активного листа

sheet.freeze_panes = 'C6'

rows = sheet.max_row
columns = sheet.max_column

header = 3

width_cols = {'Ед. изм.': 3, 'Начальный остаток': 6, 'Приход': 6, 'Расход': 7, 'Конечный остаток': 7, 'крат': 5, 'сроки': 5,
             'метка': 12, 'заяв': 7, 'разн': 7, 'без опта': 7, 'опт': 7, 'заказ в пути': 7, 'ср нов': 7, 'расчет': 7,
             'заказ филиала': 7, 'Комментарии филиала': 21, 'кон ост': 5, 'факт': 5, 'ср': 6, 'комментарии': 25, 'вес': 7}

style_cols = {'заказ филиала': green_fill, 'Комментарии филиала': green_fill}

columns_for_formulas = {'крат': None, 'расчет': None, 'вес': None}

for cell in sheet[header]:
    if cell.value:
        if cell.value in width_cols:
            sheet.column_dimensions[cell.column_letter].width = width_cols[cell.value]
        if cell.value in style_cols:
            cell.fill = style_cols[cell.value]
        if cell.value in columns_for_formulas:
            columns_for_formulas[cell.value] = cell.column_letter



for num, row in enumerate(sheet, 1):
    if row[0].value and row[0].value != 'Номенклатура':
        mult = f'{columns_for_formulas["крат"]}{num}'
        col_mult = sheet[mult].value
        if col_mult:
            requirement = f'{columns_for_formulas["расчет"]}{num}'
            sheet[f'{columns_for_formulas["вес"]}{num}'].value = f'={mult}*{requirement}'


wb.save('test.xlsx')
