from ExcelHandler import ExcelHandler
from collections import defaultdict
from Util import *

class Injector:
    def __init__(
        self,
        name,
    ):
        self.name = name
        self.items = defaultdict(dict)
    
    def add(self, item_name, color, particle, customer):
        self.items[item_name]["color"] = color
        self.items[item_name]["particle"] = particle
        self.items[item_name]["customer"] = transform_name(customer)

    def get(self):
        return self.items

class ProducibleConstraints:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "機台可生產品項說明"
        self.labels = None
        self.injector_name = []
        self.injectors = dict()
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        for i in range(self.label_row + 1, self.Excel.get_sheet_max_row(self.title)):
            row = self.Excel.get_row_value(self.title, i)
            if row[0] is None:
                continue
            
            index = row[0]
            injector_name = row[1]
            item = row[2]
            color = row[3]
            particle = row[4]
            customer = row[5]

            if injector_name not in self.injectors.keys():
                self.injectors[injector_name] = Injector(injector_name)

            self.injectors[injector_name].add(item, color, particle, customer)
            
        print("ProducibleConstraints")
        print(self.labels)
        print(self.injectors["M21"].get())
        print()