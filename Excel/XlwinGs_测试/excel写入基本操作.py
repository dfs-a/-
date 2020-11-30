import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
#创建一个新的工作簿
# wb = xw.Book('../Excel_Test/demo_01.xlsx')

#实例化工作表对象
# sht = wb.sheets['Sheet1']
# sht.range('A1').value = 'Foo 1'
# print(sht.range('A1').value)

"""range(范围扩大)----->扩大范围就是扩大一行去写入数据"""
# sht.range('A1').value = [['Foo 1','Foo 2','Foo 3'],[10.0,20.0,30.0]]
# print(sht.range('A1').expand().value)


"""功能强大的转换器能处理大多数感兴趣的数据类型"""
# df = pd.DataFrame([[1,2],[3,4]],columns=['a','b'])
# #将定义好的数据写入到这个excel表格中
# sht.range('A1').value = df
# print(sht.range('A1').options(pd.DataFrame, expand='table').value)


"""Matplotlib数字可以在Excel中显示为图片"""
# fig = plt.figure()
# plt.plot([1,2,3,4,5])
# sht.pictures.add(fig,name='MyPlot',update=True)


"""活动工作表的捷径 ： xw.Range

如果要快速与活动工作簿中的活动工作表通信，则不需要实例化工作簿和工作表对象，可以简单地执行"""
# xw.Range('A1').value = 'Foo'
# print(xw.Range('A1').value)
"""注意：在与Excel交互时，应该只使用xw.Range"""

@xw.func
def hello(name):
    return 'Hello {}'.format(name)


"""
连接到工作簿的最简单方法是由xw.Book提供：它在所有应用程序实例中查找该工作簿并返回错误，如果同一工作簿在多个实例中打开。 要连接到活动应用程序实例中的工作簿，请使用“xw.books”并引用特定应用程序
"""

# 创建app逻辑  应用-->工作簿-->工作表-->范围
# 或类似现有应用程序的xw.apps [10559]，通过xw.apps.keys（）获取可用的PID
# 创建应用
#  add_book的作用是不打开两张表，只打开一张表，因为默认xlwings操作excel的时候是否新增一个文件，默认为True
# visible参数默认为True，作用为是否看到excel操作过程
app = xw.App(add_book=False,visible=True)
# print(xw.apps.keys())
# 工作簿
wb = app.books.add()
# 工作表
sht = wb.sheets["sheet1"]
# 范围
sht.range("A1").value = "冬风诉"
# 保存excel 方式一:指定单元格来写入
sht.range("b3").value = "树华"

# 方式二:直接写入一行
sht.range("c4").value = [1,2,3,4,5,6]
# 等效于
sht.range('c4:f4').value = [5,6,7,8]

"""值得注意的是，插入列的时候向上面这样不行,因为默认是横着插入"""
"""所以:需要用到options方法指定列插入"""
sht.range("b7:b9").options(transpose=True).value = [9,10,11]


"""插入行列,采用二维数组来实现"""
sht.range("C12").value = [[2,3],[4,5]]

wb.save("demo_02.xlsx")
# 关闭excel程序
wb.close()
app.quit()

