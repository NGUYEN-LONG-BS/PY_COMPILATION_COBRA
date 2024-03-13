# Bài 11: Tương tác với Worksheet và Worksheet hiện hành
# File Excel B11_Sample.xlsx nằm trong folder chứa file script Python

import xlwings as xw

wb1 = xw.Book('B11_Sample.xlsx')    # Mở File có tên là B11_Sample.xlsx lên

#  1. Phân biệt các sheet với sheet hiện hành

# 2. Kết nối tới sheet hiện hành
WS_1 = wb1.sheets.active
print(WS_1.name, "==> tên của WS_1")   # in ra tên của sheet hiện hành

# 3. Kết nối với sheet theo chỉ số index của sheet
tong_so_sheet = wb1.sheets
print(tong_so_sheet.count, "==> File Excel có bao nhiêu sheet")

WS_2 = wb1.sheets[0]
print(WS_2.name, "==> Tên của WS_2")

WS_3 = wb1.sheets[1]
print(WS_3.name, "==> Tên của WS_3")

WS_4 = wb1.sheets[3]
print(WS_4.name, "==> Tên của WS_4")

# 4. Kết nối với sheet theo tên của sheet
WS_5 = wb1.sheets['sheet1']
print(WS_5.name, "==> Tên của WS_5")

# 5. Kết nối với sheet theo code-name của sheet



