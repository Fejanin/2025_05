import PyPDF2
import re
from datetime import datetime
import glob
import os


# TODO - не создается файл с ошибками, если в директории уже
# есть переименованный файл и есть на него непереименованный дубль!!!



def create_new_name(old_file_name: str) -> list:
    text_first_page, text_middle_pages, text_last_page = read_text(old_file_name)
    try:
        number = find_invoice_name(text_first_page)
    except:
        number = 'NO000000-00'
    try:
        max_data = find_data(text_first_page)
    except:
        max_data = '00.00.0000'
    try:
        name_driver = find_name_driver(text_last_page)
    except:
        name_driver = 'NO_NAME'
    try:
        name_company = find_name_company(text_first_page)
    except:
        name_company = 'NO_COMPANY'
    try:
        weight = find_weight(text_last_page)
    except:
        weight = '0.0'
    try:
        city = find_city(text_first_page)
    except:
        city = 'NO_CITY'
    new_file_name = f'{name_company} {number} {max_data} {city} {weight}кг ({name_driver}).pdf'
    all_sku = find_order_sku(text_first_page + text_middle_pages + text_last_page)
    all_sku = '\n'.join([str(i) for i in all_sku])
    return [old_file_name, new_file_name], [f'{new_file_name}\n{all_sku}\n\n', name_driver]


def write_report(text, name_driver):
    with open(f'{name_driver}.txt', 'a', encoding='utf-8') as f:
        f.write(text)


def find_order_sku(text):
    data = []
    pattern1 = r'SU[0-9]{6}[а-яё]{2}' # № СКЮ
    # pattern2 = r'(?<=[0-9] )[0-9]+,00 [0-9]+\.[0-9]+\.?[0-9]*' # коробки, вес брутто
    pattern2 = r'(?<=[0-9] )[0-9]+?\s?[0-9]*,00 [0-9]+\.[0-9]+\.?[0-9]*' # коробки, вес брутто
    pattern3 = r'\bра[0-9\s]+,[0-9]+' # штуки/вес нетто
    res1 = re.findall(pattern1, text)
    res2 = re.findall(pattern2, text)
    res3 = re.findall(pattern3, text)
    res2 = [i.split() for i in res2]
    for i in range(len(res2)):
        if len(res2[i]) > 2:
            res2[i] = [res2[i][0] + res2[i][1], res2[i][2]]
    res3 = [re.sub(r'\s', '_', i[2:]) for i in res3]
    data.extend(list(zip(res1, res2, res3)))
    return data


def find_weight(text):
    pattern = r'(?<=Отпуск груза произвел)[0-9\s]*'
    return ''.join(re.search(pattern, text)[0].split())


def find_invoice_name(text):
    pattern = r'[a-zA-Zа-яА-ЯёЁ]{2,}\u00ad?\d{5,}\u00ad\d{2}'
    return re.findall(pattern, text)[0].replace('\u00ad', '-')


def find_data(text):
    pattern = r'\d{2}\.\d{2}\.\d{4}'
    dates = re.findall(pattern, text)
    return '.'.join(list(reversed(str(max([datetime.strptime(i, '%d.%m.%Y') for i in dates]))[:10].split('-'))))


def find_pdf_files() -> list:
    return glob.glob('AxaptaReport*.[pP][dD][fF]')


def find_name_driver(text: str) -> str:
    pattern = r'(?<=Выданной : )[а-яА-ЯёЁ]*'
    return re.search(pattern, text)[0]


def find_city(text: str) -> str:
    cities = ["Сочи", "Орёл", "Крым", "Донецк", "Трояны", "Луганск", "Мелитополь"]
    for c in cities:
        if c in text:
            if c == "Трояны":
                c = "Бердянск"
            return c
    return "NO_CITY"


def find_name_company(text: str) -> str:
    pattern = r'ООО\s\"[А-Яа-яЁё\s]+\"'
    res = re.search(pattern, text)[0].replace('\xa0', ' ').replace('"', '')
    return res


def check_unique(data: list) -> list:
    all_files = {}
    error = {}
    for i in data:
        if not i[1] in all_files:
            all_files[i[1]] = i[0]
        else:
            td = error.get(i[1], [all_files[i[1]]])
            td.append(i[0])
            error[i[1]] = td
    return error


def rename_file(old_name: str, new_name: str) -> None:
    os.rename(old_name, new_name)


def create_error_report(error: dict) -> None:
    with open('ERROR.txt', 'w') as f:
        f.write(f'Были удалены следующие файлы:\n\n')
        for key, value in error.items():
            f.write(f'{key}:\n{value}\n\n')

def read_text(old_file_name: str) -> list:
    with open(old_file_name, 'rb') as pdfFileObj:
        pdfReader = PyPDF2.PdfReader(pdfFileObj)
        max_page = list(pdfReader.pages)
        middle = []
        for i in range(len(max_page)):
            pageObj = pdfReader.pages[i]
            if i == 0:
                start = pageObj.extract_text()
            elif i == len(max_page) - 1:
                end = pageObj.extract_text()
            else:
                middle.append(pageObj.extract_text())
        if middle:
            middle = '\n\n#new page#\n\n'.join(middle)
        else:
            middle = ''
    return start, middle, end

def find_double(lst: list) -> tuple:
    unique = {f[0][1]: [f[0][0], f[1]] for f in lst}
    del_files = [f[0] for f in lst if (f[0][1] in unique and f[0][0] != unique[f[0][1]][0])]
    for key in unique:
        write_report(*unique[key][1])
    unique = [[unique[key][0], key] for key in unique]
    return unique, del_files

def del_double(lst: list) -> None:
    res = [f[0] for f in lst]
    for f in res:
        os.remove(f)


if __name__ == '__main__':
    pdf_files = find_pdf_files()
    old_new_files = [create_new_name(i) for i in pdf_files]
    unique, del_files = find_double(old_new_files)
    error = check_unique(del_files)
    del_double(del_files)
    not_rename = [item for sublist in error.values() for item in sublist]
    for old, new in unique:
        rename_file(old, new)
    if error:
        create_error_report(error)


