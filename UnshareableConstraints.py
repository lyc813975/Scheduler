from ExcelHandler import ExcelHandler
from Util import *

class UnsharerableConstraint:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel

        self.title = "不可共用限制"
        self.constraints = []
        self.load()

    def load(self):
        for i in range(self.Excel.get_sheet_max_row(self.title)):
            row = self.Excel.get_row_value(self.title, i)
            self.constraints.append(row)
        
        print("UnsharerableConstraint")
        print(self.constraints)
        print()
            