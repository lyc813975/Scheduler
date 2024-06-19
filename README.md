# 產學案 - 排程系統

## 打包程式

1. Install pyinstaller

    ``` sh
    pip3 install pyinstaller
    ```

2. 產生exe執行檔

    ``` sh
    pyinstaller -F [target python file]
    ```

## ExcelHandler

### Dependency

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

### Install

``` sh
pip3 install openpyxl
```

### test

``` sh
python3 ExcelHandler.py
```

### excel 架構

- worksheet (excel檔案)
  - sheet1 (分頁，有title, row_height, column_width...等屬性)
    - cell A1 (有font, fill, color...等屬性)
    - cell B1 (column: A, B, ...)
    - cell A2 (row: 1, 2, ...)
    - ...
  - sheet2
    - ...

### Write Example

``` python
from ExcelHandler import ExcelHandler
from openpyxl.styles import PatternFill, Font

writer = ExcelHandler()
file = "test_handler.xlsx"
# created with at least one worksheet
writer.create_worksheet()
# add an new sheet (two sheets for now)
writer.add_sheet()
# get title of each sheet
titles = writer.get_sheet_titles()

# set cell value
writer.set_cell_value(title=titles[0], row=0, col=2, value=3)  
# set cell font (字型 大小 粗體 斜體 底線 顏色...)
writer.set_cell_font(title=titles[0], row=0, col=2, font=Font("新細明體", sz=12, bold=False, italic=False, underline=None, color='FFFFC000'))
# set cell fill (底色)
writer.set_cell_fill(title=titles[0], row=0, col=2, fill=PatternFill(fill_type="solid", fgColor="FF46A2D5"))

# save file
writer.save_worksheet(file)
```

- Result

<img src="images/example.png" alt="result" width="400"/>

### Read Example

``` python
from ExcelHandler import ExcelHandler, get_color_value

reader = ExcelHandler()
file = "test_handler.xlsx"
# load file
reader.load(file)
# get title of each sheet
titles = reader.get_sheet_titles()

value = reader.get_cell_value(title=titles[0], row=0, col=2)
font = reader.get_cell_font(title=titles[0], row=0, col=2)
fill = reader.get_cell_fill(title=titles[0], row=0, col=2)

print(f"Value: {value}, Text Color: {get_color_value(font.color)}, Background color: {get_color_value(fill.fgColor)}")
print(f"Font: {font.name}, Size: {font.size}, Bold: {font.bold}, Italic: {font.italic}, Underline: {font.underline}")
```

- Result

``` sh
Value: 3, Text Color: FFFFC000, Background color: FF46A2D5
Font: 新細明體, Size: 12.0, Bold: False, Italic: False, Underline: None
```

### TODO

- Border
