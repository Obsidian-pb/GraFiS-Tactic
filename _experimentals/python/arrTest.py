#%%
import numpy as np
import matplotlib.pyplot as plt


# %%
arr = np.loadtxt("ToMatrixTest.csv", delimiter=",")
print(arr.shape)
plt.imshow(arr)


#%%
#============Настройки========
cellsWeight = [
    [pow(2,0.5), 1, pow(2,0.5)],
    [1,          0,         1],
    [pow(2,0.5), 1, pow(2,0.5)]
]

#%%
#============Классы===================
class Claster(object):

    def __init__(self, matrix):
        self.cells = []
        self.front = []
        self.space=matrix       #1 - стена, 0 - открытое пространство

        #Матрица открытого пространства (пригодного для расчета распространения)
        vals = matrix.copy()
        vals.fill(0)
        self.vals = vals

        self.weightProfile = [
        [pow(2,0.5), 1, pow(2,0.5)],
        [1,          0, 1],
        [pow(2,0.5), 1, pow(2,0.5)]
        ]
        
    def setWeightProfile(self, weightProfile):
        '''
        Устанавливаем профиль весов клеток
        '''
        self.weightProfile = weightProfile

    def setStartPoint(self, xy):
        '''
        Устанавливаем стартовую клетку кластера
        '''
        self.cells.append(xy)
        self.front.append(xy)

    def oneStep(self, curLen):
        '''
        Один шаг расчета
        '''
        #Перебираем все клетки во фронте распротстранения
        cellsTemp = self.front.copy()
        for cell in cellsTemp:
            cy, cx = cell
            # Если длина пути текущей ячейки меньше либо равна текущему расчетному, 
            # то производим ее расчет
            if self.vals[cy,cx]<=curLen:
                # Для всех клеток в окрестности вычисляем длину пути к ним от текущей
                for dx in range(-1,2):
                    for dy in range(-1,2):
                        w = self.vals[cy,cx]+self.weightProfile[dy+1][dx+1]
                        x = cx+dx
                        y = cy+dy
                        # Если лежит в пределах матрицы
                        if self.isCellInMatrix((y,x)):
                            # Если не стена (или неучетное пространство)
                            if self.space[y,x]==0:
                                if (self.vals[y,x]>w) or (self.vals[y,x]==0):
                                    self.vals[y,x]=w
                                    if not((y,x) in self.front):
                                        self.front.append((y,x))
                                    if not((y,x) in self.cells):
                                        self.cells.append((y,x))
                # Удаляем из фронта клеток    
                self.front.remove((cy,cx))

    def getFront(self, kind='matrix'):
        '''
        Возвращаем матрицу с отмеченными клетками фронта 
        распространения расчета
        '''
        if kind=='matrix':
            matrix = self.vals.copy()
            matrix.fill(0)
            for cell in self.front:
                cy, cx = cell
                matrix[cy,cx]=2

            return matrix
        elif kind=='list':
            return self.front

    def getArea(self, kind='only'):
        area = self.vals
        for x in range(area.shape[0]):
            for y in range(area.shape[1]):
                area[x,y] = round(area[x,y],1)

        if kind=='only':
            return area
        elif kind=='inSpace':
            m = area.max()
            area = area/m
            area = area*255

            for x in range(area.shape[0]):
                for y in range(area.shape[1]):
                    area[x,y] = area[x,y] + self.space[x,y]*255*2
            return area

    def isCellInMatrix(self, xy):
        x, y = xy
        r = (x in range(self.vals.shape[0])) and (y in range(self.vals.shape[1]))
        return (r)






#%%
# arr = np.zeros((1000,1000))

yx_0 = (105, 35)
claster = Claster(arr)
claster.setStartPoint(yx_0)


# циклы расчетов
for i in range (150):
    claster.oneStep(i)
    if i%25==0:
        area = claster.getArea(kind='inSpace')
        area[(yx_0[0], yx_0[1])]=1000
        # plt.imshow(area)
        # plt.show()
        # plt.imshow(claster.getFront())
        # plt.show()

        front = claster.getFront()
        for x in range(area.shape[0]):
            for y in range(area.shape[1]):
                area[x,y] = area[x,y] + front[x,y]*255*3
        
        plt.imshow(area)
        plt.show()

#%%
plt.imshow(arr)
plt.show()

#%%
print(*[1,2,3,4,5])