from ExcelHandler import ExcelHandler

class ProducibleConstrains:
    def __init__(
        self,
        Excel: ExcelHandler
    ):
        self.Excel = Excel
        self.label_row = 1

        self.title = "機台可生產品項說明"
        self.labels = None
        self.load()

    def load(self):
        self.labels = [self.Excel.get_cell_value(self.title, self.label_row, i) for i in range(self.Excel.get_sheet_max_column(self.title))]
        print(self.labels)