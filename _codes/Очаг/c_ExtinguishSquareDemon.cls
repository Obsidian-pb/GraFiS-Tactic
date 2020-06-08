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

Dim extSquareObj As c_ExtinguishingSquare       'Объект площади тушения

Const diag As Double = 1.4142135623731         'модификатор распространения вверх и вниз
Const orto As Double = 1                      'модификатор распротсранения по диагонали
Const mmInInch As Double = 25.4

'--------------------------------Демон построитель зон площади тушения-------------------------------------------




'Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
Public Sub CreateMatrix(ByRef fireMatrix As c_Matrix)
'Активируем матрицу в соответсвии с переданными размерами
    ReDim matrix(fireMatrix.DimensionX, fireMatrix.DimensionY)
    ReDim calculatedCells(fireMatrix.DimensionX, fireMatrix.DimensionY)
    
    'Сохраняем ссылку на матрицу пожара, для дальнейшего использования
    Set fireMatrixObject = fireMatrix
End Sub

Public Property Let setGrain(ByVal grainValue)
    grain = grainValue
End Property




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare)
'Основная процедура запускающая демона - на вход передается объект площади тушения
'Демон используя волновой алгоритм рассчитывает зону площади тушения пожара
Dim block As c_ExtinguishingBlock

    'Сохраняем ссылку на объект площади тушения
    Set extSquareObj = extSquare
     
    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
        If block.alreadyCalculated = False Then
            FillExtSquare block
            block.alreadyCalculated = True
        End If
    Next block

End Sub

Public Sub RunDemonTemp(ByRef extSquare As c_ExtinguishingSquare)
'Основная процедура запускающая демона - на вход передается объект площади тушения
'Демон используя волновой алгоритм рассчитывает зону площади тушения пожара
Dim block As c_ExtinguishingBlock
Dim i As Integer
    
    'Сохраняем ссылку на объект площади тушения
    Set extSquareObj = extSquare
    
     i = 0
    'Перебираем все блоки и для каждого из них строим зоны горения
    For i = 0 To 50
        For Each block In extSquare.extinguishingBlocks
            If block.alreadyCalculated = False Then
                FillExtSquareOneStep block, i
'                block.alreadyCalculated = True
            End If
        Next block
        
        'Очищаем матрицы клеток
'        ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
'        ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))
    Next i
    

End Sub

Public Sub DrawFireSquare(ByRef extSquare As c_ExtinguishingSquare)
'Отрисовываем площадь тушения
Dim block As c_ExtinguishingBlock
Dim cell As c_Cell
Dim shp As Visio.Shape
Dim mainShp As Visio.Shape
Const mmInInch As Double = 25.4
Dim tmpColl As Collection

    
    
    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
        Set tmpColl = New Collection
        Application.ActiveWindow.DeselectAll
        
        If block.extinguishingCells.Count > 0 Then
            For Each cell In block.extinguishingCells
                If block.IsCellInThisPerimeter(cell) Then
                    Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
                            (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
                    tmpColl.Add shp
                End If
                
                
                
    '            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("ExtSquare")
    '            shp.Cells("LinePattern").FormulaU = "0"
    '            shp.Cells("FillForegnd").FormulaU = "4"
    
    '            If mainShp Is Nothing Then Set mainShp = shp
    '            Application.ActiveWindow.Select mainShp, visSelect
    '            Application.ActiveWindow.Selection.Union
    '            Set mainShp = Application.ActiveWindow.Selection(1)
    '            DoEvents
                
            Next cell
            
            'Объединяем все фигуры клеток текущего блока в единую фигуру
            For Each shp In tmpColl
                Application.ActiveWindow.Select shp, visSelect
            Next shp
            Application.ActiveWindow.Selection.Union
            Set shp = Application.ActiveWindow.Selection(1)
            CorrectSh shp
    
    '        'Указываем полученной фигуре свойство площади тушения
    '        shp.AddNamedRow visSectionProp, "ExtSquare", visTagDefault
    '        shp.CellsSRC(visSectionProp, 0, visCustPropsValue).FormulaU = str(block.GetSquare)
        End If
    Next block

End Sub

Private Function CorrectSh(ByRef shp As Visio.Shape) As Visio.Shape
'Удаляем из полученной фигуры площади тушения лишние линии (секции Geometry)
Dim secCount As Integer
Dim secMax As Integer       'Секйия с максимальным количеством строк
Dim secMin As Integer       'Секйия с минимальным количеством строк
Dim minVal As Long
Dim maxVal As Long
Dim i As Integer
    
    secCount = getShapeGeometryCount(shp)
    
    If secCount = 2 Then        'Если фигура без внутренней площади
        shp.DeleteSection visSectionFirstComponent + Int(shp.RowCount(visSectionFirstComponent) > shp.RowCount(visSectionFirstComponent + 1)) * -1
    ElseIf secCount = 4 Then    'Если фигура с внутренней площадью
        maxVal = shp.RowCount(visSectionFirstComponent)
        minVal = maxVal
        secMax = 0
        secMin = 0
        'секция 2
        If shp.RowCount(visSectionFirstComponent + 1) > maxVal Then
            maxVal = shp.RowCount(visSectionFirstComponent + 1)
            secMax = 1
        End If
        If shp.RowCount(visSectionFirstComponent + 1) < minVal Then
            minVal = shp.RowCount(visSectionFirstComponent + 1)
            secMin = 1
        End If
        'секция 3
        If shp.RowCount(visSectionFirstComponent + 2) > maxVal Then
            maxVal = shp.RowCount(visSectionFirstComponent + 2)
            secMax = 2
        End If
        If shp.RowCount(visSectionFirstComponent + 2) < minVal Then
            minVal = shp.RowCount(visSectionFirstComponent + 2)
            secMin = 2
        End If
        'секция 4
        If shp.RowCount(visSectionFirstComponent + 3) > maxVal Then
            maxVal = shp.RowCount(visSectionFirstComponent + 3)
            secMax = 3
        End If
        If shp.RowCount(visSectionFirstComponent + 3) < minVal Then
            minVal = shp.RowCount(visSectionFirstComponent + 3)
            secMin = 3
        End If
        'Удаляем промежуточные секции
        Do While i < 4
            If shp.RowCount(visSectionFirstComponent + i) = 0 Then Exit Do  'Если такой секции нет, то значит достигнут конец перебора - выходим из цикла
            
            If shp.RowCount(visSectionFirstComponent + i) <> maxVal And shp.RowCount(visSectionFirstComponent + i) <> minVal Then
                shp.DeleteSection visSectionFirstComponent + i
                i = i - 1
            End If
            
            i = i + 1
            If i > 100 Then
                Set CorrectSh = shp
                Debug.Print "Ошибка переполнения в CorrectSh"
                Exit Function
            End If
        Loop
    End If
Set CorrectSh = shp
End Function

Private Function getShapeGeometryCount(ByRef shp As Visio.Shape) As Integer
'Возвращает количество секций Geometry в фигуре shp
Dim i As Integer

    i = 0
    Do While shp.SectionExists(visSectionFirstComponent + i, 0)
        i = i + 1
    Loop
    
getShapeGeometryCount = i
End Function

Private Sub FillExtSquare(ByRef block As c_ExtinguishingBlock)
'При помощи волнового алгоритма строим площади тушения
Dim tmpColl As Collection
Dim cell As c_Cell
Dim i As Integer

    'перебираем все клетки фронта пожара и указываем в матрице расчитанных клеток, что их обрабатывать не нужно
    SetCellsToCalculated block.fireFrontCells
    
    Set tmpColl = block.fireFrontCells
    
    'Рассчитываем площадь тушения (коллекция block.extinguishingCells)
    Do While tmpColl.Count > 0
'    Do While i * grain < block.fireWay - grain
        Set tmpColl = OneStep(tmpColl, block.fireWay, (i + 1) * grain)
        
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

Private Sub FillExtSquareOneStep(ByRef block As c_ExtinguishingBlock, ByVal i As Integer)
'При помощи волнового алгоритма строим площади тушения пошагово
Dim tmpColl As Collection
Dim cell As c_Cell
'Dim i As Integer

    'перебираем все клетки фронта пожара и указываем в матрице расчитанных клеток, что их обрабатывать не нужно
    If block.extinguishingCells.Count > 0 Then
        SetCellsToCalculated block.extinguishingCells
        
        Set tmpColl = block.extinguishingCells
    Else
        SetCellsToCalculated block.fireFrontCells
        
        Set tmpColl = block.fireFrontCells
    End If
    
    'Рассчитываем площадь тушения (коллекция block.extinguishingCells)
        Set tmpColl = OneStep(tmpColl, block.fireWay, (i + 1) * grain)
        
        'Указываем для всех клеток полученной коллекции, что они уже рассчитаны
'        SetCellsToCalculated tmpColl

        'Добавляем клетки из полученной коллекции в коллекцию клеток площади тушения блока
        block.AddToExtCellsCollection tmpColl
    
    'Очищаем матрицы клеток
'    ReDim calculatedCells(UBound(calculatedCells, 1), UBound(calculatedCells, 2))
'    ReDim matrix(UBound(matrix, 1), UBound(matrix, 2))
    
    'Следом заполняем коллекцию активных клеток
'    FillActiveCellsSquare block
    
End Sub

Private Sub FillActiveCellsSquare(ByRef block As c_ExtinguishingBlock)
'При помощи волнового алгоритма находим все активные клетки относящиеся к данному блоку
Dim tmpColl As Collection
Dim cell As c_Cell
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
Private Function OneStep(ByRef inColl As Collection, ByVal maxWay As Long, ByVal maxLenCur As Long) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
        cellPower = matrix(cell.x, cell.y)
'        'Распространяем на соседние клетки
'        AddCellToColl newColl, NextCell(cell, l), cellPower + orto * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, lu), cellPower + diag * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, u), cellPower + orto * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, ru), cellPower + diag * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, r), cellPower + orto * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, rd), cellPower + diag * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, d), cellPower + orto * grain, maxWay
'        AddCellToColl newColl, NextCell(cell, ld), cellPower + diag * grain, maxWay
        'Распространяем на соседние клетки
        AddCellToColl newColl, NextCell(cell, l), cellPower + orto * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, lu), cellPower + diag * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, u), cellPower + orto * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, ru), cellPower + diag * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, r), cellPower + orto * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, rd), cellPower + diag * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, d), cellPower + orto * grain, maxWay, maxLenCur
        AddCellToColl newColl, NextCell(cell, ld), cellPower + diag * grain, maxWay, maxLenCur


    Next cell
    
    Set OneStep = newColl
    
End Function

Private Sub AddCellToColl(ByRef coll As Collection, ByRef cell As c_Cell, ByVal wayLen As Single, ByVal maxWay As Long, ByVal maxLenCur As Long)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
Dim shp As Visio.Shape

    'Если клетка находится за пределами матрицы  -выходим
    If Not IsInMatrix(cell) Then Exit Sub
    
    'Проверяем, горит ли ячейка на максимум
    If Not fireMatrixObject.IsCellMaxFire(cell.x, cell.y) Then Exit Sub
    
'Проверяем не является ли клетка клеткой стены
If fireMatrixObject.IsWall(cell.x, cell.y) Then Exit Sub
    
    If calculatedCells(cell.x, cell.y) = 1 And matrix(cell.x, cell.y) >= wayLen Then
        calculatedCells(cell.x, cell.y) = 0
        'Удаляем из имеющихся блоков
'        extSquareObj.DeleteCellFromEveryBlock cell.x, cell.y
    End If
    
    'Если путь превышает предельное значение или текущее максимальное значение или уже рассчитвана - не учитываем его
'    If wayLen >= maxWay Or wayLen > maxLenCur Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub
    If wayLen >= maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub

    'Если больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) >= wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(cell.x, cell.y) = wayLen

'                If cell.y = 0 Then
'                    Debug.Print "cell(" & cell.x & ":" & cell.y & ") wayLen=" & matrix(cell.x, cell.y)
'                End If
'                If cell.y < 15 Then
'                    Set shp = Application.ActivePage.DrawRectangle((cell.x * grain) / mmInInch, (cell.y * grain) / mmInInch, _
'                                                         (cell.x * grain + grain) / mmInInch, (cell.y * grain + grain) / mmInInch)
'                                                         shp.Cells("FillForegnd").Formula = 4
'                                                         shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("fire")
'                                                         shp.AddSection visSectionProp
'                                                         shp.AddRow visSectionProp, 0, 0
'                                                         shp.CellsSRC(visSectionProp, 0, 1).FormulaU = """Путь"""
'                                                         shp.CellsSRC(visSectionProp, 0, 0).Formula = matrix(cell.x, cell.y)
'                End If


        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
'    matrix(cell.x, cell.y) = wayLen
''    'Если путь превышает предельное значение - не учитываем его
''    If calculatedCells(cell.x, cell.y) = 1 Then Exit Sub
''
''    'Если больше предлагаемого пути, то...
''    If matrix(cell.x, cell.y) = 0 Then
''        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
''        matrix(cell.x, cell.y) = wayLen
''
''        On Error Resume Next
''        coll.Add cell, cell.x & ":" & cell.y
''    End If
End Sub

'Эти варианты для коллекции активных клеток (за пределами зоны горения):
Private Function OneStep2(ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim cell As c_Cell
Dim newColl As Collection
Dim cellPower As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
        cellPower = matrix(cell.x, cell.y)
        'Распространяем на соседние клетки
        AddCellToColl2 newColl, NextCell(cell, l), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, lu), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, u), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, ru), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, r), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, rd), cellPower + diag * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, d), cellPower + orto * grain, maxWay
        AddCellToColl2 newColl, NextCell(cell, ld), cellPower + diag * grain, maxWay

    Next cell
    
    Set OneStep2 = newColl
    
End Function

Private Sub AddCellToColl2(ByRef coll As Collection, ByRef cell As c_Cell, ByVal wayLen As Single, ByVal maxWay As Integer)
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если клетка находится за пределами матрицы  -выходим
    If Not IsInMatrix(cell) Then Exit Sub
    
    'Проверяем, является ли клетка активной, но не горит на максимум
    If Not fireMatrixObject.IsOuterActiveCell(cell.x, cell.y) Then Exit Sub
    
    'Если путь превышает предельное значение - не учитываем его
    If wayLen > maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub
    
    'Если больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(cell.x, cell.y) = wayLen
        
        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
    
End Sub

Private Sub SetCellsToCalculated(ByRef coll As Collection)
'Указываем для коллекции входящих клеток, что они уже обсчитанны
Dim cell As c_Cell
    For Each cell In coll
        calculatedCells(cell.x, cell.y) = 1
    Next cell
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'Добавляем все клетки из коллекции 2 в коллекцию 1
Dim cell As c_Cell

    On Error Resume Next
    
    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub




Private Function NextCell(ByRef cell As c_Cell, ByVal Dir As Directions) As c_Cell
'Возвращает следующую клетку согласно указанного направления
    Select Case Dir
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

Private Function IsInMatrix(ByRef cell As c_Cell) As Boolean
'Возвращает Истина, если клетка находится в пределах матрицы
    If cell.x < 0 Or cell.y < 0 Or cell.x > UBound(matrix, 1) Or cell.y > UBound(matrix, 2) Then
        IsInMatrix = False
    Else
        IsInMatrix = True
    End If
End Function
