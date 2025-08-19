from lib.funcs import check_file_name
from lib.objects import Division, ORDER, REPORT


if __name__ == "__main__":
    berdiansk = Division(
        check_file_name("Бердянск"),
        name="Бердянск")
    #doneck = None
    #lugansk = None
    #melitopol = None
    doneck = Division(
        check_file_name("Донецк"),
        name="Донецк")
    lugansk = Division(
        check_file_name("Луганск"),
        name="Луганск")
    melitopol = Division(
        check_file_name("Мелитополь"),
        name="Мелитополь")
    
    print(f"{berdiansk.not_in_db = }")  # нет в БД

    divisions = [i for i in (berdiansk, doneck, lugansk, melitopol) if i]
    
    order = ORDER("Бланк заказов сыр  дистр.xlsx", divisions)
    for d in divisions:
        print(f'{d.name = }')
        for sku in d.orders:
            if sku.in_order:
                print(f'\t{sku.name} перенесено в заказ.')
            else:
                print(f'\n\tВ заказ не занесено - {sku.name}\n')

    report = REPORT(divisions)
    print('\nREPORT:')
    for d in report.divisions:
        print(f'{d.name = }')
        print(f'{d.errors = }')
        print(f'{d.not_in_db = }')
