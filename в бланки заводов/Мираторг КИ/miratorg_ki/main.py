from my_lib.xlsx_worker import SKU, read_xlsx


# бланк заовода
FIRST_ROW = 2
ARTICLE = "A"
NAME_SKU = "B"
BOX_WEIGHT = "G"
# TODO если UNIT_MEASUREMENTS => кг, то возвращать ошибку (еще не было, проверить формулы в бланке завода)
UNIT_MEASUREMENTS = "E"
WEIGHT_SKU = "F"
ORDER = "P"  # колонка куда заносим заказ (в коробках)

sku_to_order = {}
report = ""
errors = {"Изменилось название продукции в бланке завода:\n": [],
          "Задвоенный артикул в бланке завода:\n": [],
          "Заполнены не все поля в SKU:\n": [],}

# file_data = "Ташкент 15,12.xlsx"
file_data = input("Введите название файла с заказом: ")
wb, sh = read_xlsx(file_data)

for s, o in sh:
    if o.value:
        sku_to_order[s.value.strip()] = SKU(s.value.strip(), o.value)

file_1c = "my_lib/1C.xlsx"
wb_1c, sh_1c = read_xlsx(file_1c)

for num, data in enumerate(sh_1c):
    if num == 0:  # пропускаем заголовок
        continue
    n_1c, art, n_b = data
    name = n_1c.value.strip()
    if name in sku_to_order:
        sku_to_order[name].article = art.value
        sku_to_order[name].name_manufacture = n_b.value

file_order = "новый_бланк/Бланк заказа Моспродторг (заказ на ).xlsx"
wb_order, sh_order = read_xlsx(file_order)

for row, data in enumerate(sh_order, 1):
    if row < FIRST_ROW:  # пропускаем заголовок
        continue
    for name1C in sku_to_order:  # не оптимально (усложнять лень)
        if sku_to_order[name1C].article == sh_order[f"{ARTICLE}{row}"].value:
            if sku_to_order[name1C].flag:
                # проверяем, что название СКЮ в базе совпадает с названием из бланка завода
                if sku_to_order[name1C].name_manufacture != sh_order[f"{NAME_SKU}{row}"].value:
                    errors["Изменилось название продукции в бланке завода:\n"].append(
                        f"Артикул - {sku_to_order[name1C].article}; Строка - {row}.\n\t" +
                        f"Старое название - {sku_to_order[name1C].name_manufacture}; Новое название - \
{sh_order[f"{NAME_SKU}{row}"].value}.\n"
                    )
                else:
                    # заполнить необходимые поля в SKU
                    sku_to_order[name1C].box_weight = sh_order[f"{BOX_WEIGHT}{row}"].value
                    sku_to_order[name1C].unit_maesurements = sh_order[f"{UNIT_MEASUREMENTS}{row}"].value
                    # если UNIT_MEASUREMENTS => кг, то возвращать ошибку
                    if sku_to_order[name1C].unit_maesurements != "Шт":
                        print(f"ОШИБКА. Единицы измерения не шт.!!! Строка - {row}")
                        input("Нахмите ENTER, чтобы завершить программу.")
                        raise TypeError(
                            f"ОШИБКА. Единицы измерения не шт.!!! Строка - {row}"
                        )
                    sku_to_order[name1C].weight_sku = 1 if sh_order[f"{UNIT_MEASUREMENTS}{row}"].value == "кг" \
                        else sh_order[f"{WEIGHT_SKU}{row}"].value
                    sku_to_order[name1C].num_row = row
                    sku_to_order[name1C].flag = False  # не допускаем дублирование (повторение артикула в бланке завода)
            else:  # дубль в бланке завод (создаем отчет об ошибке)
                # создать отчет об ошибке, при повторной попытке внести данные в SKU
                errors["Задвоенный артикул в бланке завода:\n"].append(
                    f"{sku_to_order[name1C].article}\n" +
                    f"\tДанные перенесены в строку {sku_to_order[name1C].num_row}." +
                    f"Название - {sku_to_order[name1C].name_manufacture}.\n" +
                    f"\tДубль находится в строке - {row}.\n"
                )

# расчитать заказы по скю
for s in sku_to_order:
    sku_to_order[s].calculate_order()


# проверка на заполнение всех полей и перенос данных в бланк завода
# если есть скю с незаполненными полями, добавляем их в errors
# если все нормально добавляем значение в REPORT.txt
for s in sku_to_order:
    res = sku_to_order[s].check_data()
    if res:
        errors["Заполнены не все поля в SKU:\n"].append(
            f"Артикул: {sku_to_order[s].article}.\n" +
            f"Наименование из 1С: {sku_to_order[s].name_1c}.\n" +
            f"Наименование из бланка завода: {sku_to_order[s].name_manufacture}.\n" +
            f"Незаполненные поля: {', '.join(res)}.\n\n"
        )
    else:  # перносим значение в новый бланк и добавляем запись в report
        sh_order[f"{ORDER}{sku_to_order[s].num_row}"] = sku_to_order[s].res_order
        text = f"{sku_to_order[s].res_order}\t{sku_to_order[s].calculate_weight()}\n"
        report += f"{sku_to_order[s].name_1c}\t{text.replace('.', ',')}"

wb_order.save("Бланк заказа Моспродторг (заказ на ).xlsx")

# создание файла с ошибками
if any([errors["Изменилось название продукции в бланке завода:\n"],
       errors["Задвоенный артикул в бланке завода:\n"],
       errors["Заполнены не все поля в SKU:\n"]]):
    with open("ERRORS.txt", "w") as f:
        for i in errors:
            if errors[i]:
                f.write(f"\n{'#'*50}\n{i}\n")
                for j in errors[i]:
                    f.write(j)

# создание файла с отчетом
with open("REPORT.txt", "w") as f:
    f.write(report)

# print(*map(str, sku_to_order.values()), sep="")
