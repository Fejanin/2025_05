from openpyxl import load_workbook


class SKU:
    def __init__(self, name_1c, order):
        # из бланка расчета
        self.name_1c = name_1c
        self.order = order
        # из бланка 1С
        self.article = None
        self.name_manufacture = None
        # из бланка завода
        self.flag = True  # проверка на повторное занесение значения (внесение данных разрешено)
        self.box_weight = None
        self.unit_maesurements = None  # если UNIT_MEASUREMENTS => кг, то WEIGHT_SKU = 1 !!!
        self.weight_sku = None
        self.num_row = None  # номер строки в бланке завода
        self.res_order = None  # вычмсляемое значение (calculate_order)

    def calculate_order(self):
        print(self.name_1c)
        try:
            weight = self.order * self.weight_sku
        except:
            print(self.check_data())
            raise TypeError
        res = round(round(weight / self.box_weight) * self.box_weight / self.weight_sku, 4)
        self.res_order = res

    def calculate_weight(self):
        return round(self.res_order * self.weight_sku, 4)

    def check_data(self):
        # возвращает список не заполенных полей
        test_data = {
            "name_1c": self.name_1c,
            "order": self.order,
            "article": self.article,
            "name_manufacture": self.name_manufacture,
            "box_weight": self.box_weight,
            "unit_maesurements": self.unit_maesurements,
            "weight_sku": self.weight_sku,
            "res_order": self.res_order
        }
        return [i for i in test_data if test_data[i] is None]

    def __str__(self):
        return f"{self.name_1c}: {self.order} => {self.article} => {self.name_manufacture}\n"

def read_xlsx(file):
    wb = load_workbook(file)
    sh = wb.active
    return wb, sh
