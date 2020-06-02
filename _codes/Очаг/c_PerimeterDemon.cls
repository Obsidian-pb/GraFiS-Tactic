VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PerimeterDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant
Private dimensionByX As Long                                        'Размерность матрицы по ширине (x)
Private dimensionByY As Long                                        'Размерность матрицы по высоте (y)

Public grain As Integer

Private perimeterCells As Collection                                'Коллекция точек периметра

'Const diag As Double = 1.4142135623731         'модификатор распространения вверх и вниз
'Const orto As Double = 1                      'модификатор распротсранения по диагонали
Const mmInInch As Double = 25.4

'--------------------------------Демон построитель периметра (площади пожара)-------------------------------------------

Private Enum CellType
    notFirePrimeter = 0
    firePerimeterNotInBlock = 1
    firePerimeterInBlock = 2
End Enum

Public Function CreateMatrix(ByVal x As Integer, ByVal y As Integer) As c_PerimeterDemon
'Активируем матрицу в соответсвии с переданными размерами
    dimensionByX = x
    dimensionByY = y
    ReDim matrix(x, y)
    
    Set CreateMatrix = Me
End Function

Public Function SetStartState(ByRef perimeterCellsCollection As Collection) As c_PerimeterDemon
'Устанавливаем стартовое состояние матрицы
    '1 Задаем стартовое состояние матрицы
    For Each cell In perimeterCellsCollection
        matrix(cell.x, cell.y) = CellType.firePerimeterNotInBlock
    Next cell
    
    Set SetStartState = Me
End Function

Public Function setGrain(ByVal gr As Integer) As c_PerimeterDemon
    grain = gr
    
    Set setGrain = Me
End Function


Public Sub RunDemon(ByRef perimeterCellsCollection As Collection)
''Разбиваем фронт пожара на блоки
'Dim block As c_ExtinguishingBlock
'
'    '1 проверяем, не входит ли точка контакта зоны работы ствола в фрагмент фронта пожара, который уже распределен
'    Set block = CheckBlockExists(extSquare, nozzle.fireContactCell)
'    If block Is Nothing Then
'        '2 разбить коллекцию фронта на участки не связанные между собой
'        BreakFront extSquare, frontCellsCollection, nozzle
'    Else
'        '3 Добавляем данные в существующий блок
'        block.expenseFact = block.expenseFact + nozzle.waterExpense
'    End If
    
    
'Строим последовательность клеток
Dim cell As c_Cell
    
    SetStartState perimeterCellsCollection
    
    Set perimeterCells = New Collection
    Set cell = perimeterCellsCollection(1)
    
    TryAddCell perimeterCells, cell
'    For Each cell In perimeterCellsCollection
    
    

End Sub




'Private Sub BreakFront(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection, _
'                    ByRef nozzle As c_FireNozzleModel)
''Разбиваем фронт пожара на участки
''Dim cell As c_Cell
'Dim extBlock As c_ExtinguishingBlock
'Dim tmpColl As Collection
'
'    'Перебираем все клетки в коллекции стартовых клеток
''    For Each cell In startCellsCollection
'
'        'если клетка еще не распределена по блокам, начинаем от нее рекурсивный поиск всех клеток площади тушения данного блока
''        If matrix(startCell.x, startCell.y) = CellType.fireFrontNotInBlock Then
'            'создаем новую коллекцию
'            Set tmpColl = New Collection
'            'Для новой коллекции выполняем рекурсивное наполнение от текущей ячейки
'            TryAddCell tmpColl, nozzle.fireContactCell
'
'            'Создаем новый блок и передаем ему временную коллекцию клеток фронта
'            Set extBlock = New c_ExtinguishingBlock
'            Set extBlock.fireFrontCells = tmpColl
'            extBlock.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
'            extBlock.AddToExtCellsCollection tmpColl
''            AddCellsToCollection extBlock.extinguishingCells, tmpColl
'            '---Устанавливаем глубину тушения
'                If nozzle.nozzleType = waterHand Then
'                    extBlock.fireWay = 5000     'Для ручных стволов
'                Else
'                    extBlock.fireWay = 10000     'Для лафетных стволов
'                End If
'                extBlock.expenseFact = extBlock.expenseFact + nozzle.waterExpense
'                extBlock.grain = grain
'
'            'Добавляем созданный блок в коллекцию блоков модели площади тушения
'            extSquare.AddExtinguishingBlock extBlock
''        End If
''    Next cell
'
'End Sub



Private Sub TryAddCell(ByRef thisColl As Collection, ByVal cell As c_Cell)
'Пытаемся добавить клетку в коллекцию (рекурсия)
    
    'Если указанная точка находится за пределами матрицы - выходим из процедуры
    If cell.x < 0 Or cell.y < 0 Or cell.x > dimensionByX - 1 Or cell.y > dimensionByY - 1 Then Exit Sub
    
    'Проверяем является ли данная клетка клеткой периметра пожара
    If matrix(cell.x, cell.y) = CellType.notFirePrimeter Then Exit Sub
    
    'Проверяем, нет ли клетки уже в коллекции
    If matrix(cell.x, cell.y) = CellType.firePerimeterInBlock Then Exit Sub
    
    'Добавляем клетку в коллецию
    AddCellToCollection thisColl, cell
    matrix(cell.x, cell.y) = CellType.firePerimeterInBlock
    
    'Если нет - рекурсия на следующие клетки
        TryAddCell thisColl, NextCell(cell, l)
'        TryAddCell thisColl, NextCell(cell, lu)
        TryAddCell thisColl, NextCell(cell, u)
'        TryAddCell thisColl, NextCell(cell, ru)
        TryAddCell thisColl, NextCell(cell, r)
'        TryAddCell thisColl, NextCell(cell, rd)
        TryAddCell thisColl, NextCell(cell, d)
'        TryAddCell thisColl, NextCell(cell, ld)
   
End Sub

Private Sub AddCellToCollection(ByRef thisColl As Collection, ByVal cell As c_Cell)
'Добавляем в указанную коллекцию указанную клетку
    thisColl.Add cell, cell.x & ":" & cell.y
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'Добавляем все клетки из коллекции 2 в коллекцию 1
Dim cell As c_Cell

    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub

'Private Function CheckBlockExists(ByRef extSquare As c_ExtinguishingSquare, ByRef cell As c_Cell) As c_ExtinguishingBlock
''Пытаемся найти блок к которому может относиться пришедшая клетка. Еслитн аходим возвращаем ссылку на него, иначе - Nothing
'Dim blck As c_ExtinguishingBlock
'
'    For Each blck In extSquare.extinguishingBlocks
'        If blck.IsCellInThisFront(cell) Then
'            Set CheckBlockExists = blck
'            Exit Function
'        End If
'    Next blck
'Set CheckBlockExists = Nothing
'End Function





Private Function NextCell(ByRef cell As c_Cell, ByVal dir As Directions) As c_Cell
'Возвращает следующую клетку согласно указанного направления
    Select Case dir
        Case Is = Directions.l
            Set NextCell = NewCell(cell.x - 1, cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(cell.x - 1, cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(cell.x, cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(cell.x + 1, cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(cell.x + 1, cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(cell.x + 1, cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(cell.x, cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(cell.x - 1, cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    Set NewCell = cell
End Function




'-----------Собственно отрисовка площади------------
Public Sub DrawSquare()
'Отрисовываем площадь используя для этого построение DrawPoiyLine
Dim cell As c_Cell
'Dim arr As Variant
    
Dim shp As Visio.Shape
Dim perimterPntsStr As String
Dim perimterPnts() As Double
Dim intCounter As Integer
Dim i As Long

'    'Инициализируем массив с координатами периметра
'    For Each cell In perimeterCells
'        perimterPntsStr = perimterPntsStr & str(cell.x / 25.4) & ";" & str(cell.y / 25.4) & ";"
'    Next cell
'    perimterPntsStr = Left(perimterPntsStr, Len(perimterPntsStr) - 1)
'    perimterPnts = Split(perimterPntsStr, ";")
'
'    Set shp = ActivePage.DrawPolyline(perimterPnts, 0)
    
    
    i = 0
    ReDim perimterPnts(perimeterCells.Count * 2 - 1)
    'Инициализируем массив с координатами периметра
    For Each cell In perimeterCells
        perimterPnts(i) = cell.x * grain / 25.4
        i = i + 1
        perimterPnts(i) = cell.y * grain / 25.4
        i = i + 1
    Next cell
    
    Set shp = ActivePage.DrawPolyline(perimterPnts, 0)
    
    
    
'    adblXYPoints(1) = 100
'    adblXYPoints(2) = 100
'    adblXYPoints(3) = 300
'    adblXYPoints(4) = 300
'    adblXYPoints(5) = 500
'    adblXYPoints(6) = 100
'    adblXYPoints(7) = 100
'    adblXYPoints(8) = 200
'
'    'Use the DrawPolyline method to draw a shape that has 2-D behavior.
'    Set vsoShape = ActivePage.DrawPolyline(adblXYPoints, 0)

''Increase the Y-coordinate of the array by 4 to separate
''the next shape drawn from the first.
'For intCounter = 2 To UBound(adblXYPoints) Step 2
'    adblXYPoints(intCounter) = adblXYPoints(intCounter) + 4
'Next intCounter
'
''Use the DrawPolyline method to draw a shape that has 1-D behavior.
'Set vsoShape = ActivePage.DrawPolyline(adblXYPoints, visPolyline1D)

    
End Sub


