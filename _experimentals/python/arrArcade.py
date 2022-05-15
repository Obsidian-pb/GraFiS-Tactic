#%%
import numpy as np
import arcade

#%%
arr = np.loadtxt("ToMatrixTest.csv", delimiter=",")

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
        
    def setStartPoint(self, xy):
        self.cells.append(xy)
        self.front.append(xy)
        # self.vals[xy[0],xy[1]]=1

    def oneStep(self, curLen):
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
                        w = self.vals[cy,cx]+cellsWeight[dy+1][dx+1]
                        x = cx+dx
                        y = cy+dy
                        # Если лежит в пределах матрицы
                        if self.isCellInMatrix((y,x)):
                            # Если не стена
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

    def getArea(self):
        area = self.vals
        for x in range(area.shape[0]):
            for y in range(area.shape[1]):
                area[x,y] = round(area[x,y],1)
        return area

    def isCellInMatrix(self, xy):
        x, y = xy
        r = (x in range(self.vals.shape[0])) and (y in range(self.vals.shape[1]))
        return (r)

#%%
#%%
# arr = np.zeros((1000,1000))

yx_0 = (50, 125)
claster = Claster(arr)
claster.setStartPoint(yx_0)


# циклы расчетов
for i in range (200):
    claster.oneStep(i)
    print('step {}'.format(i))

area = claster.getArea()
m = area.max()
area = area/m
area = area*255
area[(yx_0[0], yx_0[1])]=255
# plt.imshow(area)