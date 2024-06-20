from ExcelHandler import ExcelHandler
from ProducibleConstrains import ProducibleConstrains

if __name__ == '__main__':
    path = "sample.xlsx"
    Excel = ExcelHandler(path)
    Excel.load()
    print(Excel.get_sheet_titles())
    Producible = ProducibleConstrains(Excel)