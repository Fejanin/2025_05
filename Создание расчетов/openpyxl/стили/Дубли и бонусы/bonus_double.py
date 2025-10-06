def get_columns_abc_by_name(name_cols, num_header_row, sheet):
    name_cols_dict = {key: None for key in name_cols}
    for cell in sheet[num_header_row]:
        if cell.value:
            if cell.value in name_cols_dict:
                name_cols_dict[cell.value] = cell.column_letter
    return name_cols_dict

def get_max_col(row_header, sheet):
    max_col = None
    for i in range(1, sheet.max_column):
        if sheet.cell(row=HEADER, column=i).value:
            max_col = i
    return max_col


def get_max_row(sheet):
    max_row = None
    for i in range(1, sheet.max_row):
        if sheet.cell(row=i, column=1).value:
            max_row = i
    return max_row


if __name__ == "__main__":
    import openpyxl
    from openpyxl.styles import PatternFill, Font

    # file_name = 'дв 26,06,25 тшрсч пок ки.xlsx'
    file_name = input('Введите название файла: ')
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active

    # константы
    HEADER = 3
    ABC_MARK2 = list(get_columns_abc_by_name(('метка2', ), HEADER, sheet).values())[0]

    name_cols = ('Расход', 'Конечный остаток', 'опт')
    yellow_fill = PatternFill(start_color='ffff00', end_color='ffff00', fill_type='solid')
    red_font = Font(color='ff0000')
    
    abc_cols_by_name = get_columns_abc_by_name(name_cols, HEADER, sheet)
    max_col = get_max_col(HEADER, sheet)
    max_row = get_max_row(sheet)

    valid_sku = {}
    not_valid_sku = {}

    ERRORS = {}

    for i in range(HEADER + 1, max_row + 1):
        if sheet.cell(i, 1).value:
            if sheet[f'{ABC_MARK2}{i}'].value:
                key_name = sheet[f'{ABC_MARK2}{i}'].value
                values = {}
                not_valid_sku.setdefault(key_name, []).append(i)
            else:
                key_name = sheet[f'A{i}'].value
                if key_name in valid_sku:
                    raise ValueError('Одинаковые значения в названии номенклатуры.')
                valid_sku[key_name] = i

    strings_res = {}  # храним итоговые строки - 'E43': '=225.879+E67'
    color_cells = []
    for i in not_valid_sku:
        if i not in valid_sku:
            ERRORS.setdefault('ERROR_KEY_ORIGIN', []).append((i, not_valid_sku[i]))
        else:
            for row in not_valid_sku[i]:
                cells_valid = [f'{col}{valid_sku[i]}' for col in abc_cols_by_name.values()]
                cells_not_valid = [f'{col}{row}' for col in abc_cols_by_name.values()]
                for v, not_v in zip(cells_valid, cells_not_valid):
                    if sheet[not_v].value:
                        if not v in strings_res:
                            num_v = sheet[v].value or 0  # сохраняем начальное значение
                            strings_res[v] = f'={num_v}+{not_v}'
                            color_cells.append(v)
                            color_cells.append(not_v)
                        else:
                            strings_res[v] += f'+{not_v}'
                            color_cells.append(not_v)

    for address in strings_res:
        sheet[address].value = strings_res[address]

    for cell in color_cells:
        sheet[cell].fill = yellow_fill
        sheet[cell].font = red_font
    
    if ERRORS:
        with open('ERROR.txt', 'w+') as f:
            f.write(str(ERRORS))

    wb.save(f'res_{file_name}')
    
