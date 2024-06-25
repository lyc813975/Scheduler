from ExcelHandler import ExcelHandler
from collections import defaultdict
from Util import *
import copy

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
        self.products = defaultdict(list)
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        injectors = []
        products = []
        colors = []
        particles = []
        for i in range(self.label_row + 1, self.Excel.get_sheet_max_row(self.title)):
            row = self.Excel.get_row_value(self.title, i)
            injectors.append(row[1])
            products.append(row[2])
            colors.append(row[3])
            particles.append(row[4])
        print("讀取 labels")
        print(self.labels)
        print()

        print("讀取射出機欄位")
        print(fill_merge_cell(injectors))
        