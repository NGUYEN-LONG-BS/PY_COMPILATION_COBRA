# Bài 08: Làm việc với nhiều Workbook

import xlwings as xw    # import thu vien xlwings

# Khoi tao cac Workbook moi
xw.App(add_book=False)    # co tham so False, nen se khong tao them WB
xw.App()
xw.Book()

# Gan bien
wb1 = xw.books['book1']
wb2 = xw.books['book2']

# Kiem tra gan bien co thanh cong khong
print(wb1.name)
print(wb2.name)
print("Gán biến thành công")

# Doi ten cua WB dang mo
wb1.save('book1_14h09.xlsx')
wb2.save('book2_14h09.xlsx')

# Gan lai bien
wb1 = xw.books['book1_14h09.xlsx']
wb2 = xw.books['book2_14h09.xlsx']

# Dong cac WB lai
wb1.close()
wb2.close()

# /