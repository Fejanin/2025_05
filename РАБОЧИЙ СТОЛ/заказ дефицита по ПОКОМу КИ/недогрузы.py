import os
from openpyxl import load_workbook


def get_all_files():
    all_files = []
    for root, dirs, files in os.walk("."):
            for file in files:
                file_path = str(os.path.join(root, file))
                all_files.append(file_path)
    return all_files


def get_xlsx_files(files):
    return [f for f in files if ".xlsx" in f.lower()]


def get_data_from_file(file):
    data = []
    wb = load_workbook(file)
    sh = wb.active

    cols = {
        "Код единицы продаж": None,
        "Название": None,
        "Недогруз/урезание, кг": None,
        }

    for r in sh:
        for num, i in enumerate(r):
            if i.value in cols:
                cols[i.value] = num
        break
    for r in sh:
        if "SU" in r[cols["Код единицы продаж"]].value and "-" in r[cols["Недогруз/урезание, кг"]].value:
            data.append(
                (r[cols["Код единицы продаж"]].value.strip(),
                r[cols["Название"]].value.strip(),
                str_to_num(r[cols["Недогруз/урезание, кг"]].value.strip()),)
                )
    return data


def str_to_num(s):
    n = "".join(s.split())
    return int(n)


if __name__ == "__main__":
    files = get_all_files()
    files = get_xlsx_files(files)
    sku = get_data_from_file(files[0])
    branches = {
        "Бердянск": {},
        "Донецк": {},
        "Луганск": {},
        "Мелитополь": {},
        }
    total = {}
    for br in branches:
        for f in files:
            if br in f:
                #print(f)
                data = get_data_from_file(f)
                #print(data)
                for i in data:
                    sku = i[0]
                    name = i[1]
                    num = i[2]
                    if sku in branches[br]:
                        branches[br][sku][1] += num
                    else:
                        branches[br][sku] = [name, num]
                    if sku in total:
                        total[sku][1] += num
                    else:
                        total[sku] = [name, num]
    #print(branches)
    for br in branches:
        with open(f"{br}.txt", "w") as f:
            for key in branches[br]:
                f.write(f"{key}\t{branches[br][key][0]}\t{branches[br][key][1]}\n")
    with open(f"TOTAL.txt", "w") as f:
        for key in total:
            f.write(f"{key}\t{total[key][0]}\t{total[key][1]}\n")







