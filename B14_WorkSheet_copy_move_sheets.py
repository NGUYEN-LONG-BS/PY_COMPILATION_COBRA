# Bài 14 Copy Worksheet, di chuyển Worksheet sang vị trí mới

import xlwings as xw
wb = xw.Book(r'C:\Users\ADMIN\Desktop\hoc_python\B14_Sample.xlsx')

# Copy sheet
# lưu ý: wb.sheets ==> crlt + click chuột trái vào chữ sheet để mở file main.py
# wb.sheets['Sheet1'].copy(name='sheet mới được copy')
# wb.sheets['Sheet1'].copy(name='python', after=wb.sheets('Sheet1'))

# di chuyển sheet sang vị trí mới (vị trí cuối cùng)
# vì không tìm thấy thuộc tính move, nên làm theo 3 bước sau

# wb.sheets['Sheet1'].copy()      # Tên mặt định sẽ là Sheet1 (2)
# wb.sheets['Sheet1'].delete()
wb.sheets['Sheet1 (2)'].name = 'Sheet1'