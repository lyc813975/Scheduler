from ExcelHandler import ExcelHandler
from ProducibleConstraints import ProducibleConstraints
from ColdMoldConstraints import ColdMoldConstraints
from HotMoldConstraints import HotMoldConstraints
from UnshareableConstraints import UnsharerableConstraint

if __name__ == '__main__':
    path = "sample.xlsx"
    Excel = ExcelHandler(path)
    Excel.load()
    Producible = ProducibleConstraints(Excel)
    Cold = ColdMoldConstraints(Excel)
    Hot = HotMoldConstraints(Excel)
    Unshare = UnsharerableConstraint(Excel)