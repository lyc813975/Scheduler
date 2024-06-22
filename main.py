from ExcelHandler import ExcelHandler
from ProducibleConstrains import ProducibleConstrains

if __name__ == '__main__':
    path = "sample.xlsx"
    Excel = ExcelHandler(path)
    Excel.load()
    Producible = ProducibleConstrains(Excel)