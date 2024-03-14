# Bài 13 Cách lấy danh sách của tất cả Worksheet
import xlwings as xw
wb = xw.Book(r'C:\Users\ADMIN\Desktop\hoc_python\B13_Sample.xlsx')

# Tạo 12 sheet, tương ứng với 12 tháng
for i in range(12):
    wb.sheets.add(f'month{i + 1}', after=wb.sheets[i])          # i bắt đầu từ số 0
    
# Khai báo list danh sách các tháng, ban đầu để list này rỗng
l_sh = []

# Tiến hành lấy danh sách tên các sheet, sử dụng phương thức append()
for sh in wb.sheets:
    l_sh.append(sh.name)

print(l_sh)
