# Bài 10: Làm việc với WB có PassWord

# File Excel ở cùng thư mục với file Script Python và có tên là B10_Pass_123.xlsx với mật khẩu là 123

import xlwings as xw

wb5 = xw.Book(fullname = r'C:\Users\ADMIN\Desktop\du_an_01\B10_Pass_123.xlsx',
              password="123",
              )
wb5.close()
