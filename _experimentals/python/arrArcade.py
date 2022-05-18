#%%
import numpy as np
import arcade
import matplotlib.pyplot as plt

#%%
arr = np.loadtxt("ToMatrixTest.csv", delimiter=",")
print(arr.shape)

#%%
#============Настройки========
cellsWeight = [
    [pow(2,0.5), 1, pow(2,0.5)],
    [1,          0, 1],
    [pow(2,0.5), 1, pow(2,0.5)]
]
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
yx_0 = (50, 125)

# arr = np.zeros((1000,1000))


claster = Claster(arr)
claster.setStartPoint(yx_0)


# циклы расчетов
for i in range (200):
    claster.oneStep(i)
    # print('step {}'.format(i))

area = claster.getArea()
m = area.max()
area = area/m
area = area*255
area[(yx_0[0], yx_0[1])]=255
#%%
plt.imshow(area)

#%%
# Задать константы для размеров экрана
SCREEN_WIDTH = arr.shape[1]
SCREEN_HEIGHT = arr.shape[0]

# Открыть окно. Задать заголовок и размеры окна (ширина и высота)
arcade.open_window(SCREEN_WIDTH, SCREEN_HEIGHT, "Drawing Example")

# Задать белый цвет фона.
# Для просмотра списка названий цветов прочитайте:
# http://arcade.academy/arcade.color.html
# Цвета также можно задавать в (красный, зеленый, синий) и
# (красный, зеленый, синий, альфа) формате.
arcade.set_background_color(arcade.color.WHITE)

# Начать процесс рендера. Это нужно сделать до команд рисования
arcade.start_render()

# Рисуем стены
# try:
for x in range(SCREEN_WIDTH):
    for y in range(SCREEN_HEIGHT):
        if arr[y, x]==1:
            arcade.draw_point(x, y, arcade.color.BLACK, size=1)
            # arcade.finish_render()
        if area[y, x]>0:
            arcade.draw_point(x, y, arcade.color.RED, size=1)
    # arcade.finish_render()
# except:
#     arcade.finish_render()

# Завершить рисование и показать результат
arcade.finish_render()

# Держать окно открытым до тех пор, пока пользователь не нажмет кнопку “закрыть”
arcade.run()

# %%
