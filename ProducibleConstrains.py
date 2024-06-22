from ExcelHandler import ExcelHandler
import copy
class ProducibleConstrains:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "機台可生產品項說明"
        self.labels = None
        self.injector_name = None
        self.load()

    def load(self):
        self.labels = self.Excel.get_row_value(self.title, self.label_row)
        
        ranges = copy.deepcopy(self.Excel.workbook[self.title].merged_cells)
        for rang in ranges:
            self.Excel.workbook[self.title].unmerge_cells(rang.coord)

        test = [self.Excel.get_cell_value(self.title, 3, i) for i in range(self.Excel.get_sheet_max_column(self.title))]
        print(test)
        print(self.labels)