from ExcelHandler import ExcelHandler
from collections import defaultdict
from Util import *
import copy

class ColdMoldConstraints:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "模具限制說明(冷模)"
        self.labels = None
        self.products = defaultdict(list)
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        # self.type = self.Excel.get_col_value(self.title, 0)
        print(self.labels)
        