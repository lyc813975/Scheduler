from ExcelHandler import ExcelHandler
from collections import defaultdict
from Util import *
import copy

class HotMoldConstraints:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "模具限制說明(熱模)"
        self.labels = None
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        print(self.labels)
        