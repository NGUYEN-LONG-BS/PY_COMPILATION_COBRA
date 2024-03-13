# Bài 7 Cách tạo Workbook mới, cách lưu và đóng Workbook

import xlwings as xw    # Thư viện chuyên dùng để xử lý Excel

#I Creat a new book, save change the name and close it
#1. Creat a new book:
# xw.Book() # cách 1: Tạo ra một file Excel mới
xw.App() # cách 2: Tạo ra một file Excel mới