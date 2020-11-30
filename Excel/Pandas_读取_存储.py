import pandas as pd
import numpy as np
import xlwings as xw
app = xw.App(add_book=False,visible=False)
wb = app.books.open("./Excel_Test/10月汉滨区住宅成交数据.xlsx")
sht1 = wb.sheets['普通住宅']
project_Name = sht1.range("b2:b80").value
data = pd.read_excel('./Excel_Test/10月汉滨区住宅成交数据.xlsx',sheet_name = "普通住宅")
# print(data)
# io = r'./Excel_Test/10月汉滨区住宅成交数据.xlsx'
# data_1 = pd.read_excel(io,sheet_name="普通住宅",usecols=[1])
# data_2 = pd.read_excel(io,sheet_name="普通住宅",usecols=[5])
print(data.info())
print(data.head(100))
Project_Name = []
for i in project_Name:
    if i not in Project_Name:
        Project_Name.append(i)

# Building = []
# for i in Project_Name:
some = data[(data['QubuilderName'] == '锦绣星城')]
# print(some)
#指定套数列求和

# Sum = some['最高价'].min()
# Sum_1 = some['最高价'][1]
Sum = some['均价']
Sum_1 = some['面积']
Sum_2 = some['面积'].sum()

for i,n in zip(Sum,Sum_1):
    print(i,n)
# Building.append(Sum)
# bu = Building
print(Sum,Sum_1,Sum_2)

#矩阵合并
# df1 = pd.DataFrame(np.ones((3,4))*0,columns=['a','b','c','d'])
# df2 = pd.DataFrame(np.ones((3,4))*1,columns=['a','b','c','d'])
# df3 = pd.DataFrame(np.ones((3,4))*2,columns=['a','b','c','d'])
#
# print(df1)
# print(df2)
# print(df3)
# #省略原矩阵的索引,全部按顺序排列索引
# res = pd.concat([df1,df2,df3],axis=0,ignore_index=True)
# print(res)