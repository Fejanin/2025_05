import openpyxl
from openpyxl.styles import PatternFill


green_fill = PatternFill(start_color='13E22E', end_color='13E22E', fill_type='solid')


# file_name = 'new_дв 19,12,24 днрсч пок зпф.xlsx'
file_name = input('Введите название файла: ')

wb = openpyxl.load_workbook(file_name)

sheet = wb.active  # Получение активного листа

sheet.freeze_panes = 'C6'

rows = sheet.max_row
columns = sheet.max_column

header = 3

width_cols = {'Ед. изм.': 3, 'Начальный остаток': 6, 'Приход': 6, 'Расход': 7, 'Конечный остаток': 7, 'крат': 5, 'сроки': 5,
             'метка': 9, 'заяв': 6.5, 'разн': 5.5, 'без опта': 6, 'опт': 6, 'заказ в пути': 6, 'ср нов': 6, 'расчет (потребность)': 7,
             'расчет (отгрузит завод)': 7, 'заказ филиала': 7, 'Комментарии филиала': 20, 'кон ост': 5, 'факт': 5, 'ср': 6, 'комментарии': 24, 
             'вес': 6, 'крат кор': 6, 'заказ кор.': 6, 'ВЕС': 6, 'ряд': 5, 'паллет': 5, 'кол-во паллет': 6}

style_cols = {'заказ филиала': green_fill, 'Комментарии филиала': green_fill}

columns_for_formulas = {'крат': None, 'расчет (потребность)': None, 'расчет (отгрузит завод)': None,
                        'вес': None, 'крат кор': None, 'заказ кор.': None, 'ВЕС': None,
                        'ряд': None, 'паллет': None, 'кол-во паллет': None}

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
            mult_boxes = f'{columns_for_formulas["крат кор"]}{num}'
            requirement = f'{columns_for_formulas["расчет (потребность)"]}{num}'
            layer = f'{columns_for_formulas["ряд"]}{num}'
            order_boxes = f'{columns_for_formulas["заказ кор."]}{num}'
            width_by_layers = f'{columns_for_formulas["ВЕС"]}{num}'
            
            sheet[f'{columns_for_formulas["вес"]}{num}'].value = f'={mult}*{requirement}'
            sheet[f'{order_boxes}'].value = f'=MROUND({requirement}, {mult_boxes}*{layer})/{mult_boxes}'
            sheet[f'{width_by_layers}'].value = f'={order_boxes}*{mult_boxes}*{mult}'
            sheet[f'{columns_for_formulas["расчет (отгрузит завод)"]}{num}'].value = f'={mult_boxes}*{order_boxes}'
            sheet[f'{columns_for_formulas["кол-во паллет"]}{num}'].value = f'={order_boxes}/{columns_for_formulas["паллет"]}{num}'

wb.save('test.xlsx')
