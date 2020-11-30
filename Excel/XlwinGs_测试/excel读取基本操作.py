import xlwings as xw
"""读excel的常用方式"""
# 操作思维逻辑
# 应用->工作簿->工作表->范围

app = xw.App(visible=True,add_book=False)

# 打开excel表格文件
wb = app.books.open("demo_02.xlsx")
# 选定工作表
sht = wb.sheets["sheet1"]

# 范围-读取某个位置的值
print(sht.range("a1").value)

# 读一行
print(sht.range("c4:h4").value)

# 读一列
print(sht.range("b7:b9").value)

# 读行列
print(sht.range("c12:d13").value)