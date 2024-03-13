# Bài 7 Cách tạo Workbook mới, cách lưu và đóng Workbook

import xlwings as xw    # Thư viện chuyên dùng để xử lý Excel

#I Creat a new book, save change the name and close it
#1. Creat a new book:
# xw.Book() # cách 1: Tạo ra một file Excel mới
xw.App() # cách 2: Tạo ra một file Excel mới

#2. Save the new book:
wb1 = xw.books.active
# print(wb1.name)
wb1.save("thenewbook1.xlsx")

#3. close the new book:
wb1.close()