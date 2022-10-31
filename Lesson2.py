import numpy as np
a = np.genfromtxt('матрица1.txt', delimiter=',', dtype= None )
print(a)
#получаем размерность матрицы my_data
n = a.shape[0]

#создаем единичную матрицу, с такой же размерностью как матрица a
e = np.eye(n)

#умножаем матрицу a на матрицу e
u = np.dot(a, e)
#сохраняем результат в файл
np.savetxt("result.txt", u, delimiter=",")
print(a)
print('----------')
print(u)
print(e)
