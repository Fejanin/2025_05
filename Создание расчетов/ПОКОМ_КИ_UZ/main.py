from my_lib.base_func import create_base_table_from_dv_file
from my_lib.base_func import add_data_from_old_file
from my_lib.base_func import add_data_from_sales_file
from my_lib.base_func import create_new_file
import os
import re


# files
file_dv = input('Введите название файла с поступлением товара: ')
file_dv_old = input('Введите название старого файла с расчетом: ')
file_initial_balance = input('Введите название файла с начальным остатком: ')
file_final_balance = input('Введите название файла с окончательными остатком: ')
file_sales = input('Введите название файла с продажами: ')
# file_declared = input('Введите название файла с заявками: ')
file_sales_wholesale = input('Введите название файла с продажами оптовым клиентам: ')


# за основу берем файл движения, извлекать только приходы и СКЮ
table = create_base_table_from_dv_file(file_dv)


# извлекаем данные из файла прошлой недели и добавляем данные в таблицу
add_data_from_old_file(file_dv_old, table)


# извлекаем начальный остаток
add_data_from_sales_file(file_initial_balance, table, ('Количество начальное', 'initial_balance'))


# извлекаем начальный остаток
add_data_from_sales_file(file_final_balance, table, ('Количество конечное', 'final_balance'))


# извлекаем данные из файла продаж
add_data_from_sales_file(file_sales, table, ('Итого', 'sales'))


# извлекаем данные из файла заявлено
# add_data_from_sales_file(file_declared, table, ('Заказано', 'declared'))


# извлекаем данные из файла с заказами оптовых клиентов
if file_sales_wholesale:
    add_data_from_sales_file(file_sales_wholesale, table, ('Количество', 'wholesale'))


# создаем новый файл и заносим в него данные
create_new_file(f'new_{file_dv}', table)

'''
print(len(table.get_rows()))
for i in table.rows:
    print(i.get_columns())
'''
