import numpy as np

#定义矩阵类型
array = np.array([[1,2,3],[4,5,6]],dtype=np.int)
print(array)

#查看具体为几维数组
print('number of dim:',array.ndim)
#查看具体有几行几列
print('shape:',array.shape)
#查看矩阵大小
print('size:',array.size)
#生成三行四列的矩阵,数据内容为零
print(np.zeros((3,4)))
#生成范围内的数组
print(np.arange(10,20,2))

#生成范围内的三行四列的的矩阵
print(np.arange(12).reshape((3,4)))
A = np.arange(12).reshape((3,4))
#匹配出矩阵中最小值的索引
print(np.argmin(A))
#匹配矩阵中最da值索引
print(np.argmax(A))

#计算整个矩阵的平均值
print(np.mean(A))

#取出矩阵的中位数
print(np.median(A))

#生成范围内段长的一维数组
print(np.linspace(1,10,5))

#生成原矩阵相同的累加矩阵
print(np.cumsum(A))

#生成原矩阵每两个之间的差
print(np.diff(A))

#生成原矩阵不是零的坐标，---索引
print(np.nonzero(A))

#将矩阵的行列进行调换--转置
print(np.transpose(A))
#运算矩阵之间的乘法
print((A.T).dot(A))

#小于5的全部变成5大于9的全部变成9
print(np.clip(A, 5, 9))

#axis等于0对列进行计算取中间平均值，等于1对行进行取中间平均值计算
print(np.mean(A,axis=0))


B = np.array([1,1,1,1])
C = np.array([2,2,2,2])
print(np.vstack((B,C)))  #上下合并矩阵

print(np.hstack((B,C)))  #左右合并矩阵

print(np.concatenate((C,B,B,C),axis=0)) #多行合并 0=行  1=列