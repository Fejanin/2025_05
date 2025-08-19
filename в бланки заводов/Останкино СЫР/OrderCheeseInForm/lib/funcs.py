import os


def check_file_name(name):
    while True:
        file_name = input(f"Введите название файла для филиала {name}: ")
        if not file_name:
            print(f"Заказ для {name} делать не нужно.")
            return None
        if os.path.isfile(file_name):
            print(f"Файл для {name} найден.")
            return file_name
        print(f"Файла с названием <<{file_name}>> нет в данной директории.\nПопробуйте еще раз.\n")


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
        if sheet.cell(row=row_header, column=i).value:
            max_col = i
    return max_col


def get_max_row(sheet):
    max_row = None
    for i in range(1, sheet.max_row):
        if sheet.cell(row=i, column=1).value:
            max_row = i
    return max_row
