from ExcelHandler import ExcelHandler
from collections import defaultdict
from Util import *
import copy

class HotMold:
    def __init__(
        self,
        name,
    ):
        self.name = name
        self.items = dict()
    
    def add(self, mold_number, machine_number, n_cavities, time):
        self.items["mold_number"] = mold_number
        self.items["machine_number"] = machine_number
        self.items["n_cavities"] = n_cavities
        self.items["time"] = time

    def get(self):
        return self.items

class HotMoldConstraints:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "模具限制說明(熱模)"
        self.labels = None
        self.products = dict()
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        for i in range(self.label_row + 1, self.Excel.get_sheet_max_row(self.title)):
            row = self.Excel.get_row_value(self.title, i)
            if row[0] is None:
                continue
            
            mold_number = row[0]
            time = row[1]
            type = row[2]
            n_cavities = row[3]
            machine_number = row[4]
            time = row[5]

            if type not in self.products.keys():
                self.products[type] = HotMold(type)

            self.products[type].add(mold_number, machine_number, n_cavities, time)

        print("HotMoldConstraints")
        print(self.labels)
        print(self.products["MAX-385"].get())
        