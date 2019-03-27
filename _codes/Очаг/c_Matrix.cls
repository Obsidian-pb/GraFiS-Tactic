VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public fuelValueLayer As Variant                            'Запас горючего - текущий. Если равен 0 , клетка гаснет и больше не загорается.
'Public lowerBurnBoundLayer As Variant                       'Нижняя граница при которой клетка начинает распространять горение на соседние (а так же вообще брабатывается)
'Public maximumBurnPowerLayer As Variant                     'Максимальная мощность горения (после которой увеличение мощности прекращается)

Private dimensionByX As Long                                        'Размерность матрицы по ширине (x)
Private dimensionByY As Long                                        'Размерность матрицы по высоте (y)

Private openSpaceLayer As Variant                               'Матрица ограждающих конструкций
Private currentgPowerLayer As Variant                        'Текущая мощность горнеия (в пределе от 0 до maximumBurnPower) запас горючего
Private gettedPowerInOneStepLayer As Variant                 'Энергия полученная клеткой за один шаг расчета (в конце шага суммируется с текущей мощностью)


Const cellPowerModificator As Double = 1
Const lowerBurnBound As Double = 7                          'Нижняя граница при которой клетка начинает распространять горение на соседние (а так же вообще брабатывается)
Const maximumBurnPower As Double = 100                      'Максимальная мощность горения (после которой увеличение мощности прекращается)

Public firedCells As Collection                             'Коллекция горящих клеток
Public fireMaxCells As Collection                           'Коллекция горящих на максимум клеток (собственно зона горения)
Public chargedCells As Collection                           'Коллекция клеток получившех в данном шаге энергию
Public startedFireCells As Collection                       'Коллекция клеток, которые только загораются, но их мощность еще не максимальна
Public activeFiredCells As Collection                       'Коллекция активных горящих клеток - тех, что участвуют в расчете - позволяет ускорить расчет в тактической модели, но в физической неприменимо, т.к. на позволяет уитывать гаснущие при тушении клетки
Public fireFrontCells As Collection                         'Коллекция клеток фронта пожара - клеток горящих по максимуму и находящихся на границе зоны распространения огня
Public keyCornerCells As Collection                         'Коллекция ключевых клеток - клеток в которых есть и огонь и стена и не поностью горящие клетки


'---------------------процедуры создания матрицы------------------------------
Public Sub CreateMatrix(ByVal dimX As Long, dimY As Long)
'Инициируем матрицу:
    dimensionByX = dimX
    dimensionByY = dimY
    
    '---Инициируем слои в соответствии с переданным размером матрицы
    ReDim currentgPowerLayer(dimX, dimY) As Double
    ReDim gettedPowerInOneStepLayer(dimX, dimY) As Double
    ReDim openSpaceLayer(dimX, dimY) As Byte
    
    '---Инициируем коллекции клеток
    Set firedCells = New Collection
    Set fireMaxCells = New Collection
    Set chargedCells = New Collection
    Set startedFireCells = New Collection
    Set activeFiredCells = New Collection
    Set fireFrontCells = New Collection
    Set keyCornerCells = New Collection
End Sub

Public Sub SetOpenSpace(ByRef outerMatrix() As Variant)
'Заполняем слой ограждающих конструкций
Dim i As Long
Dim j As Long
    
    For i = 0 To dimensionByX
        For j = 0 To dimensionByY
            openSpaceLayer(i, j) = outerMatrix(i, j)
        Next j
    Next i
    
End Sub








Public Sub SetCellValue(ByVal x As Long, ByVal y As Long, ByVal layer As MatrixLayerType, ByVal value As Variant)
'Устанавливаем значение для клетки с указанными координатами в указанном слое
    On Error GoTo EX

    Select Case layer
        Case Is = MatrixLayerType.mtrOpenSpaceLayer
            openSpaceLayer(x, y) = value
        Case Is = MatrixLayerType.mtrCurrentgPowerLayer
            currentgPowerLayer(x, y) = value
        Case Is = MatrixLayerType.mtrGettedPowerInOneStepLayer
            gettedPowerInOneStepLayer(x, y) = value
    End Select
Exit Sub
EX:
    Debug.Print "Error: x=" & x & ", y=" & y & ", layer=" & layer & ", value=" & value
End Sub

Public Sub SetFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'Создаем новую точку горения
    
    'Если указанная точка находится за пределами матрицы - выходим из процедуры
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then Exit Sub
    
    'Если указанная точка является стеной - выходим из процедуры
    If openSpaceLayer(x, y) = CellState.csWall Then Exit Sub
    
    'Указываем данные новой точки
    If value = 0 Then
        currentgPowerLayer(x, y) = maximumBurnPower
        'добавляем в коллекцию фронта пожара
        fireFrontCells.Add CreateNewCell(x, y), x & ":" & y
        'добавляем в коллекцию площади пожара
        fireMaxCells.Add CreateNewCell(x, y), x & ":" & y
    Else
        currentgPowerLayer(x, y) = value
        activeFiredCells.Add CreateNewCell(x, y), x & ":" & y
    End If
    
    'Добавляем запись о горящей точке в коллекцию
    firedCells.Add CreateNewCell(x, y)
End Sub





'-----------------------------То что связано с распространением горения--------------------------------
Public Sub RizeCellFire(ByVal x As Long, ByVal y As Long, ByVal value As Double)
'Увеличиваемтекущую мощность горения данной клетки
    currentgPowerLayer(x, y) = currentgPowerLayer(x, y) + value
    If currentgPowerLayer(x, y) > maximumBurnPower Then currentgPowerLayer(x, y) = maximumBurnPower
End Sub



'-------------------------------Свойства--------------------------------------------
Public Property Get DimensionX() As Long
    DimensionX = dimensionByX
End Property
Public Property Get DimensionY() As Long
    DimensionY = dimensionByY
End Property
Public Property Get GetMaximumBurnPower() As Long
    GetMaximumBurnPower = maximumBurnPower
End Property


Public Function GetFirePower(ByVal x As Long, ByVal y As Long) As Double
'Возвращаем мощность горения клетки, при этом, если она ниже порогового значения распространения, возвращаем ноль
Dim tmpFirePower As Double
    
    tmpFirePower = currentgPowerLayer(x, y)
    GetFirePower = IIf(tmpFirePower < lowerBurnBound, 0, tmpFirePower)
    
End Function
Public Function GetActualFirePower(ByVal x As Long, ByVal y As Long) As Double
'В отличие от GetFirePower возвращаем реальную мощность горения клетки, даже если она ниже порогового значения распространения
    GetActualFirePower = currentgPowerLayer(x, y)
End Function


Public Function GetPowerInOneStep(ByVal x As Long, ByVal y As Long) As Double
    GetPowerInOneStep = gettedPowerInOneStepLayer(x, y)
End Function


Public Function GetOpenSpaceLayerValue(ByVal x As Long, ByVal y As Long) As Byte
    GetOpenSpaceLayerValue = openSpaceLayer(x, y)
End Function


Public Function IsInner(ByVal x As Long, ByVal y As Long) As Boolean
'Возвращет Истина, если вокруг клетки только горящие (или стены), Ложь, если есть хоть одна не полностью горящая
    IsInner = True
    'по диагонали
    If IsCellCanFire(x - 1, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x - 1, y + 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y + 1) Then
        IsInner = False
        Exit Function
    End If
    'по ортогонали
    If IsCellCanFire(x, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x, y + 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x - 1, y) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y) Then
        IsInner = False
        Exit Function
    End If
End Function

Public Function IsTripleStateCell(ByVal x As Long, ByVal y As Long) As Boolean
'Возвращает Истина, если ячейка находится на границе трех сред - огня, стен и свободного пространства (имеет све их в окружении)
'Поскольку рассматривается и так горящая клетка, на горение прилегающиен е проверяются
Dim havingWall As Boolean
Dim havingOpenSpace As Boolean
    
    'Проверяем наличие стен
    If IsCellsAroundExists(x, y, 1) Then havingWall = True
    
    'Проверяем не является ли клетка внутренней
    havingOpenSpace = Not IsInner(x, y)
    
    IsTripleStateCell = havingWall And havingOpenSpace

End Function


Private Function IsCellsAroundExists(ByVal x As Long, ByVal y As Long, ByVal value As Integer) As Boolean
'Возвращаем Истина, если клетки вокруг, в слое ограждающих конструкций имеют указанное значение (хоть одна), иначе, Ложь
    'по диагонали
    If IsCellsAreType(x - 1, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x - 1, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    'по ортогонали
    If IsCellsAreType(x, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x - 1, y, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
IsCellsAroundExists = False
End Function
Private Function IsCellsAreType(ByVal x As Long, ByVal y As Long, ByVal value As Integer) As Boolean
'Проверяем какого типа данная ячейка (для слоя ограждающих конструкций)
    'Проверяем, входит ли клетка в пределы матрицы
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
        IsCellsAreType = False
        Exit Function
    End If
    
    IsCellsAreType = openSpaceLayer(x, y) = value
End Function

'Private Function IsCellsAroundNotMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
''Проверяем, имеются ли в окружении клетки не горящие на полную мощность
'    'по диагонали
'    If IsCellNotMaxFire(x - 1, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x - 1, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    'по ортогонали
'    If IsCellNotMaxFire(x, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x - 1, y) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'IsCellsAroundNotMaxFire = False
'End Function
'Private Function IsCellNotMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
''Возвращает Истина, если клетка горит не на полную мощность (либо вообще не горит) и Ложь, если нет
'    'Проверяем, входит ли клетка в пределы матрицы
'    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
'        IsCellNotMaxFire = False
'        Exit Function
'    End If
'
'    IsCellNotMaxFire = currentgPowerLayer(x, y) < maximumBurnPower
'End Function

Private Function IsCellCanFire(ByVal x As Long, ByVal y As Long) As Boolean
'Возвращает Истина, если клетка может гореть и Ложь, если нет
    
    'Проверяем, входит ли клетка в пределы матрицы
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
        IsCellCanFire = False
        Exit Function
    End If
    
    'Проверяем, не является ли клетка стеной
    If openSpaceLayer(x, y) = CellState.csWall Then
        IsCellCanFire = False
        Exit Function
    End If
    
    'Проверяем мощность горения клетки
    If currentgPowerLayer(x, y) < maximumBurnPower Then
        IsCellCanFire = True
    Else
        IsCellCanFire = False
    End If
    
End Function

Public Function IsWall(ByVal x As Long, ByVal y As Long) As Boolean
'Проверяем, находится ли в указанной клетке стена
    On Error GoTo EX
    If openSpaceLayer(x, y) = 1 Then
        IsWall = True
    Else
        IsWall = False
    End If
Exit Function
EX:
    IsWall = True
End Function

Public Function IsCellMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
'Проверяем, горит ли клетка на полную мощность
    On Error GoTo EX
    IsCellMaxFire = currentgPowerLayer(x, y) = maximumBurnPower
Exit Function
EX:
    IsCellMaxFire = False
End Function

Public Function IsOuterActiveCell(ByVal x As Long, ByVal y As Long) As Boolean
'Проверяем, является ли клетка внешней активной, т.е. еще не горит на полную мщность и находится за пределами площади горения
Dim val As Double
    
    On Error GoTo EX
    
    val = GetFirePower(x, y)
    IsOuterActiveCell = val > 0 And val < maximumBurnPower
        
Exit Function
EX:
    
End Function

'-----------------------------Инструменты-------------------------------------------
Private Function CreateNewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
Dim cell As c_Cell
    
    Set cell = New c_Cell
    Set cell.ParentMatrix = Me
    cell.x = x
    cell.y = y
    
Set CreateNewCell = cell
End Function
