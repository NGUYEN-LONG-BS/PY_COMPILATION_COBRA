# Bài 12: Tạo sheet mới, đổi tên, xoá sheet

# File Excel B12_Sample.xlsx nằm trong folder chứa file script Python

import xlwings as xw
wb1 = xw.Book('B12_Sample.xlsx')    # Mở File có tên là B12_Sample.xlsx lên

# 1. Tạo sheet mới
wb1.sheets.add()                                                            # Tự động tạo ra sheet mới, nằm ở vị trí index = 0, tên là sheet...
wb1.sheets.add('tên sheet được tạo mới', after='Sheet1')                    # Tạo mới sheet có tên: Tên sheet được tạo mới (nằm phía sau sheet1)
wb1.sheets.add('tên sheet mới', before='Sheet1')                            # Tạo mới sheet có tên: Tên sheet mới (nằm phía trước sheet1)
wb1.sheets.add('Sheet_new')                                                 # Tạo mới sheet có tên: Tên Sheet_new

# 2. Đổi tên sheet
for sh in wb1.sheets:
    if sh.name == 'Sheet_new':
        sh.name = 'Sheet_data'

# 3. Xoá sheet
for sh in wb1.sheets:
    if sh.name == 'Sheet_data':
        sh.delete()

# 4. Xoá hết các sheet, chỉ để lại 1 sheet
for sh in wb1.sheets:
    if sh.name != 'Sheet1':
        sh.delete()