class Row:
    COLOR = '46d444' # зеленый цвет для новых СКЮ
    def __new__(cls, name=None, *args, **kwargs):
        '''Создает объект класса только если передан аргумент - name.'''
        if name:
            return super().__new__(cls)

    
    def __init__(self, name):
        self.name = name # Номенклатура
        self.units_of_measurement = None # Ед. изм.
        self.initial_balance = None # Начальный остаток
        self.receipt_of_products = None # Приход
        self.sales = None # Продажи
        self.final_balance = None # Конечный остаток
        self.mark = None # Метка
        self.multiplicity = None # Кратность
        self.expiration_dates = None # Сроки годности
        self.declared = None # Заявлено | заяв
        self.diff = None # разн
        self.orders_is_on_the_way = {} # Заказы в пути {'29,11,': 50, '01,12,: 120}
        self.orders_is_on_the_way_last = {}  # Заказы в пути (в коробках)
        self.average_values = {}  # средние значения продаж => {'17,11,': 100, '24,11,': 111, ...}
        self.comments = None # комментарии
        self.boxes = None # кол-во коробок в заказе
        self.layer = None # кол-во коробок в слое
        self.pallet = None # кол-во коробок в паллете
        self.number_of_pallets = None # кол-во паллет


    def get_columns(self):
        return self.__dict__


    def __eq__(self, obj):
        if obj.__class__.__name__ == 'Row':
            return self.name == obj.name
        return self.name == obj

    def __lt__(self, obj):
        if obj.__class__.__name__ == 'Row':
            return self.name < obj.name
        return self.name < obj


    def calculate_order(self):
        for item in self.orders_is_on_the_way_last:
            if self.orders_is_on_the_way_last[item]:
                self.orders_is_on_the_way_last[item] *= self.boxes or 0


    def __str__(self):
        return self.name


class Table:
    COLUMNS = {
        # '<имя>': ['<название колонки>', <номер колонки / словарь с подкатегориями и номерами колонок>', <сумма колонки>]
        'name': ['Номенклатура', None, False],
        'units_of_measurement': ['Ед. изм.', None, False],
        'initial_balance': ['Начальный остаток', None, False],
        'receipt_of_products': ['Приход', None, False],
        'sales': ['Расход', None, True],
        'final_balance': ['Конечный остаток', None, True],
        'multiplicity': ['крат', None, False],
        'expiration_dates': ['сроки', None, False],
        'mark': ['метка', None, False],
        'declared': ['заяв', None, True],
        'diff': ['разн', None, True],
        'without_wholesale': ['без опта', None, True],
        'wholesale': ['опт', None, True],
        'orders_is_on_the_way': ['заказ в пути', {}, True],
        'orders_is_on_the_way_last': ['заказ в пути', {}, True],
        'new_average_sales': ['ср нов', None, True],
        'order_requirement': ['расчет (потребность)', None, True],
        'order_will_ship': ['расчет (отгрузит завод)', None, True],
        'order_from_f': ['заказ филиала', None, True],
        'comments_from_f': ['Комментарии филиала', None, False],
        'remains': ['кон ост', None, False],
        'fact': ['факт', None, False],
        'average_values': ['ср', {}, True],
        'comments': ['комментарии', None, False],
        'weight': ['вес', None, True],
        'boxes': ['крат кор', None, False],
        'boxes ordered': ['заказ кор.', None, True],
        'weight_boxes': ['ВЕС', None, True],
        'layer': ['ряд', None, False],
        'pallet': ['паллет', None, False],
        'number_of_pallets': ['кол-во паллет', None, True],
    }
    START_HEADER = 3
    START_ROWS = 6
    END_ROWS = 500
    ORDERS_IS_ON_THE_WAY = {}
    AVERAGE_VALUES = {}
    
    def __init__(self):
        self.rows = []
        self.date = None


    def add_row(self, row):
        self.rows.append(row)


    def get_header(self):
        return self.COLUMNS


    def get_rows(self):
        '''
        Return sorted list self.rows
        '''
        return sorted(self.rows)


    def number_columns(self):
        num_col = 1
        for i in self.COLUMNS:
            if type(self.COLUMNS[i][1]) is dict:
                for j in self.COLUMNS[i][1]:
                    self.COLUMNS[i][1][j] = num_col
                    num_col += 1
            else:
                self.COLUMNS[i][1] = num_col
                num_col += 1



    
