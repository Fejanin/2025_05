import moduls.worker_xlsx as W


file = input('Введите название файла: ')
new_file = input('Введите название бланка-завода: ') #  чистый бланк

total_weight = 0 # общий вес перенесенный в бланк заказа
good_report = []
error = {
    'double_order': [],
    'not_zero': [],
    'not_found_ost': [],
    'old_sku': [],
    'not_found_in_traslater': []
    }
reader_data = W.ReaderData(file)
error['double_order'] += reader_data.error
translater = W.Translater(r'moduls\1С.xlsx')
writer = W.Writer(new_file)
old_sku = W.OldSKU(r'moduls\old_SKU.xlsx')


for i in reader_data.all_keys:
    if i in translater.all_keys: # соответствие между позициями в заказе клиента и "переводчике"
        if translater.all_keys[i] in writer.all_keys:
            adress = f'E{writer.all_keys[translater.all_keys[i]]}'
            box = writer.ws[adress].value
            if box:
                error['not_zero'].append(f'{i} {translater.all_keys[i]} невозможно внести в бланк, т.к. ячейка {adress} уже содержит значение {box}.')
            else:
                value_order = reader_data.all_keys[i]["order"]
                writer.ws[adress] = value_order
                good_report.append(f'Заносим значение {i} {translater.all_keys[i]} в кол-ве {reader_data.all_keys[i]["order"]}\n\t\t в ==> {translater.all_keys[i][1]}, ячейка {adress}')
                total_weight += value_order
        else:
            error['not_found_ost'].append(f'{i} - {translater.all_keys[i]}/ не найден в бланке завода.')
    else:
        if i  in old_sku.all_keys:
            error['old_sku'].append(f'{i} из строки {reader_data.all_keys[i]["num_row"]} - устаревшее СКЮ.')
        else:
            error['not_found_in_traslater'].append(f'{i} из строки {reader_data.all_keys[i]["num_row"]} НЕ НАЙДЕНО в бланке 1С.xlsx (бланк-переводчик).')

writer.wb.save(new_file)



with open('REPORT.txt', 'w') as f:
    if (any([bool(error[i]) for i in error])):
        f.write(f'Обнаружены следующие ошибки:\n')
        for i in error:
            if error[i]:
                f.write('=' * 50 + '\n')
            for j in error[i]:
                f.write(j + '\n')
        f.write('#' * 50 + '\n\n')
    for i in good_report:
        f.write(i + '\n')
    f.write('\n')
    f.write(f'Общий вес в бланке заказа (не задвоенных СКЮ) - {sum([reader_data.all_keys[i]["order"] for i in reader_data.all_keys])}.\n')
    f.write(f'Вес перенесенных СКЮ в заказник завода составляет - {total_weight}.\n')


