from ExcelHandler import ExcelHandler
from Util import *

class ColdMold:
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

class ColdMoldConstraints:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "模具限制說明(冷模)"
        self.labels = None
        self.products = dict()
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        for i in range(self.label_row + 1, self.Excel.get_sheet_max_row(self.title)):
            row = self.Excel.get_row_value(self.title, i)
            if row[0] is None:
                continue

            type = row[0]
            mold_number = row[1]
            machine_number = row[2]
            n_cavities = row[3]
            time = row[4]
            _ = row[5]

            if type not in self.products.keys():
                self.products[type] = ColdMold(type)

            self.products[type].add(mold_number, machine_number, n_cavities, time)

        # self.type = self.Excel.get_col_value(self.title, 0)
        print("ColdMoldConstraints")
        print(self.labels)
        # print(self.products.keys())
        print(self.products["MPX-078"].get())
        print()
        