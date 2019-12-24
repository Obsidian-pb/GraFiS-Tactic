VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishSquareDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant                           'Матрица расчетной зоны
Dim calculatedCells As Variant                  'Матрица рассчитанных клеток
Dim grain As Integer                            'Размер зерна

Public fireMatrixObject As c_Matrix             'Матрица расчетной зоны


Const diag As Double = 1.4142135623731         'модификатор распространения вверх и вниз
Const orto As Double = 1                      'модификатор распротсранения по диагонали

'--------------------------------Демон построитель зон площади тушения-------------------------------------------




'Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
Public Sub CreateMatrix(ByRef fireMatrix As c_Matrix)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim matrix(fireMatrix.DimensionX, fireMatrix.DimensionY)
    ReDim calculatedCells(fireMatrix.DimensionX, fireMatrix.DimensionY)
    
    'Сохраняем ссылку на матрицу пожара, для дальнейшего использования
    Set fireMatrixObject = fireMatrix
End Sub

Public Property Let SetGrain(ByVal grainValue)
    grain = grainValue
End Property




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare)
'Основная процедура запускающая демона - на вход передается объект площади тушения
'Демон используя волновой алгоритм рассчитывает зону площади тушения пожара
Dim block As c_ExtinguishingBlock


     
    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
        If block.alreadyCalculated = False Then
            FillExtSquare block
            block.alreadyCalculated = True
        End If
    Next block

End Sub

Public Sub DrawFireSquare(ByRef extSquare As c_ExtinguishingSquare)
'Отрисовываем площадь тушения
Dim Cell As c_Cell
Dim shp As Visio.Shape
Const mmInInch As Double = 25.4

    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
'        FillExtSquare block

        For Each Cell In block.extinguishingCells

            Set shp = Application.ActivePage.DrawRectangle(Cell.x * grain / mmInInch, Cell.y * grain / mmInInch, _
                    (Cell.x + 1) * grain / mmInInch, (Cell.y + 1) * grain / mmInInch)
            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("ExtSquare")
            shp.Cells("LinePattern").FormulaU = "0"
            shp.Cells("FillForegnd").FormulaU = "4"
        Next Cell
    Next block

End Sub

Private Sub FillExtSquare(ByRef block As c_ExtinguishingBlock)
'При помощи волнового алгоритма строим площади тушения
Dim tmpColl As Collection
Dim Cell As c_Cell
Dim i As Integer

    'перебираем все клетки фронта пожара и указываем в матрице расчитанных клеток, что их обрабатывать не нужно
    SetCellsToCalculated block.fireFrontCells
    
    Set tmpColl = block.fireFrontCells
    
    'Рассчитываем площадь тушения (коллекция block.extinguishingCells)
    Do While tmpColl.Count > 0
        Set tmpColl = OneStep(tmpColl, block.fireWay)
        
        'Указываем для всех клеток полученной коллекции, что они уже рассчитаны
        SetCellsToCalculated tmpColl

        'Добавляем клетки из полученной коллекции в коллекцию клеток площади тушения блока
        block.AddToExtCellsCollection tmpColl
'        AddCellsToCollection block.extinguishingCells, tmpColl
        'Байпасс
        i = i + 1
        If i > 10000 Then
            Debug.Print "Аварийный выход из цикла DoWhile в процедуре FillExtSquare, класса c_ExtinguishSquareDemon"
            Exit Do
        End If
    Loop
    
    'Очищаем матрицы клеток
    ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
    ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))
    
    'Следом заполняем коллекцию активных клеток
    FillActiveCellsSquare block
    
End Sub

Private Sub FillActiveCellsSquare(ByRef block As c_ExtinguishingBlock)
'При помощи волнового алгоритма находим все активные клетки относящиеся к данному блоку
Dim tmpColl As Collection
Dim Cell As c_Cell
Dim i As Integer

    'перебираем все клетки фронта пожара и указываем в матрице расчитанных клеток, что их обрабатывать не нужно
    SetCellsToCalculated block.fireFrontCells

    Set tmpColl = block.fireFrontCells

    'Рассчитываем площадь тушения (коллекция block.activeCells)
    Do While tmpColl.Count > 0
        Set tmpColl = OneStep2(tmpColl, block.fireWay)

        'Указываем для всех клеток полученной коллекции, что они уже рассчитаны
        SetCellsToCalculated tmpColl

        'Добавляем клетки из полученной коллекции в коллекцию клеток площади тушения блока
        block.AddToActiveCellsCollection tmpColl
'        AddCellsToCollection block.activeCells, tmpColl
        'Байпасс
        i = i + 1
        If i > 10000 Then
            Debug.Print "Аварийный выход из цикла DoWhile в процедуре FillActiveCellsSquare, класса c_ExtinguishSquareDemon"
            Exit Do
        End If
    Loop

    'Очищаем матрицы клеток
    ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
    ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))

End Sub

'Эти варианты для коллекции клеток площади тушения (в зоне горения):
Private Function OneStep(ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim Cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each Cell In inColl
        cellPower = matrix(Cell.x, Cell.y)
        'Распространяем на соседние клетки
        AddCellToColl newColl, NextCell(Cell, l), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, lu), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, u), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, ru), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, r), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, rd), cellPower + diag * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, d), cellPower + orto * grain, maxWay
        AddCellToColl newColl, NextCell(Cell, ld), cellPower + diag * grain, maxWay

    Next Cell
    
    Set OneStep = newColl
    
End Function

Private Sub AddCellToColl(ByRef coll As Collection, ByRef Cell As c_Cell, ByVal wayLen As Single, ByVal maxWay As Integer)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если клетка находится за пределами матрицы  -выходим
    If Not IsInMatrix(Cell) Then Exit Sub
    
    'Проверяем, горит ли ячейка на максимум
    If Not fireMatrixObject.IsCellMaxFire(Cell.x, Cell.y) Then Exit Sub
    
    'Если путь превышает предельное значение - не учитываем его
    If wayLen > maxWay Or calculatedCells(Cell.x, Cell.y) = 1 Then Exit Sub
    
    'Если больше предлагаемого пути, то...
    If matrix(Cell.x, Cell.y) = 0 Or matrix(Cell.x, Cell.y) > wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(Cell.x, Cell.y) = wayLen

        On Error Resume Next
        coll.Add Cell, Cell.x & ":" & Cell.y
    End If
    
End Sub

'Эти варианты для коллекции активных клеток (за пределами зоны горения):
Private Function OneStep2(ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim Cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each Cell In inColl
        cellPower = matrix(Cell.x, Cell.y)
        'Распространяем на соседние клетки
        AddCellToColl2 newColl, NextCell(Cell, l), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, lu), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, u), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, ru), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, r), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, rd), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, d), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(Cell, ld), cellPower + diag * grain, maxWay

    Next Cell
    
    Set OneStep2 = newColl
    
End Function

Private Sub AddCellToColl2(ByRef coll As Collection, ByRef Cell As c_Cell, ByVal wayLen As Single, ByVal maxWay As Integer)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если клетка находится за пределами матрицы  -выходим
    If Not IsInMatrix(Cell) Then Exit Sub
    
    'Проверяем, является ли клетка активной, но не горит на максимум
    If Not fireMatrixObject.IsOuterActiveCell(Cell.x, Cell.y) Then Exit Sub
    
    'Если путь превышает предельное значение - не учитываем его
    If wayLen > maxWay Or calculatedCells(Cell.x, Cell.y) = 1 Then Exit Sub
    
    'Если больше предлагаемого пути, то...
    If matrix(Cell.x, Cell.y) = 0 Or matrix(Cell.x, Cell.y) > wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(Cell.x, Cell.y) = wayLen
        
        On Error Resume Next
        coll.Add Cell, Cell.x & ":" & Cell.y
    End If
    
End Sub

Private Sub SetCellsToCalculated(ByRef coll As Collection)
'Указываем для коллекции входящих клеток, что они уже обсчитанны
Dim Cell As c_Cell
    For Each Cell In coll
        calculatedCells(Cell.x, Cell.y) = 1
    Next Cell
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'Добавляем все клетки из коллекции 2 в коллекцию 1
Dim Cell As c_Cell

    For Each Cell In col2
        col1.Add Cell, Cell.x & ":" & Cell.y
    Next Cell
End Sub




Private Function NextCell(ByRef Cell As c_Cell, ByVal dir As Directions) As c_Cell
'Возвращает следующую клетку согласно указанного направления
    Select Case dir
        Case Is = Directions.l
            Set NextCell = NewCell(Cell.x - 1, Cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(Cell.x - 1, Cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(Cell.x, Cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(Cell.x + 1, Cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(Cell.x + 1, Cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(Cell.x + 1, Cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(Cell.x, Cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(Cell.x - 1, Cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
Dim Cell As c_Cell
    Set Cell = New c_Cell
    Cell.x = x
    Cell.y = y
    
    Set NewCell = Cell
End Function

Private Function IsInMatrix(ByRef Cell As c_Cell) As Boolean
'Возвращает Истина, если клетка находится в пределах матрицы
    If Cell.x < 0 Or Cell.y < 0 Or Cell.x > UBound(matrix, 1) Or Cell.y > UBound(matrix, 2) Then
        IsInMatrix = False
    Else
        IsInMatrix = True
    End If
End Function
