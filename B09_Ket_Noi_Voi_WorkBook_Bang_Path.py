# Bài 9 Kết nối Script Python với Workbook.

import xlwings as xw    # import thu vien xlwings

# Trường hợp 1: File scrypt Python nằm chung folder với file Excel có tên: B09_Connect.xlsx
# Cần phải cd đến đường dẫn của file Excel đó trước ==> cách này mất thời gian
wb3 = xw.Book('B09_Connect.xlsx')

# Trường hợp 2: Tương tác với file Exel bằng path
wb4 = xw.Book(r'C:\Users\ADMIN\Desktop\du_an_01\B09_Connect.xlsx')