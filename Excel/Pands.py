import pandas as pd
import numpy as np

#会创建一个索引
s = pd.Series([1,3,6,np.nan,44,1])
# print(s)

#增加值
dates = pd.date_range('20201127',periods=7)
# print(dates)

#生成表格
df = pd.DataFrame(np.random.randn(7,4),index=dates,columns=['a','b','c','d'])
# print(df)

#不设置index生成表格,默认是从0开始生成索引
df_1 = pd.DataFrame(np.arange(12).reshape((3,4)))
# print(df_1)

df2 = pd.DataFrame({
    'A':1,
    'B':pd.Timestamp('20201127'),
    'C':pd.Series(1,index=list(range(4)),dtype='int32'),
    'D':np.array([3] * 4,dtype='int32'),
    'E':pd.Categorical(['test','train','test','train']),
    'F':'foo'
})
print(df2)
#查看矩阵类型
# print(df2.dtypes)
#查看index
# print(df2.index)
#查看列的名字
# print(df2.columns)
#查看每一行的数据
# print(df2.values)
#运算矩阵里面数字的平均值
# print(df2.describe())
#矩阵颠倒
# print(df2.T)

#将列反向排序
print(df2.sort_index(axis=1,ascending=False))
#将行反向排序
print(df2.sort_index(axis=0, ascending=False))
#对指定的值排序
print(df2.sort_values(by='E'))


dates_1 = pd.date_range('20201128',periods=6)
df_2 = pd.DataFrame(np.arange(24).reshape((6,4)),index=dates_1,columns=['A','B','C','D'])

#指定索引改值
df_2.iloc[2,2] = 1234
print(df_2)
#指定位置改值
df_2.loc['20201128','A'] = 3333
print(df_2)
#指定范围改值
df_2[df_2.A>4] = 1
print(df_2)

#增加列
df_2['F'] = np.nan
print(df_2)
df_2['E'] = pd.Series([1,2,3,4,5,6],index=pd.date_range('20201128',periods=6))
print(df_2)


#处理数据丢失情况
df_2[0,1] = np.nan
df_2[1,2] = np.nan
print(df.dropna(axis=0,how='any')) #如果how为any就是只要有一个为null就删除这一行  ‘all’是全部为null就删除


#将为null的值填充为指定的值
print(df_2.fillna(value=0))

#查看哪些数据已丢失
print(df_2.isnull())

#查看是否丢失数据
print(np.any(df_2.isnull()) == True)