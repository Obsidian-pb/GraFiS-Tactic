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
Private conturPoints As Collection                                  'Коллекция точек контура

Public fireMatrixObject As c_Matrix             'Матрица расчетной зоны

Dim extSquareObj As c_ExtinguishingSquare       'Объект площади тушения

Const diag As Double = 1.4142135623731         'модификатор распространения вверх и вниз
Const orto As Double = 1                      'модификатор распротсранения по диагонали
Const mmInInch As Double = 25.4

'--------------------------------Демон построитель зон площади тушения-------------------------------------------
Private Sub Class_Terminate()
    Set conturPoints = Nothing
End Sub



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
Dim cell As c_Cell
Dim i As Integer

    'Сохраняем ссылку на объект площади тушения
    Set extSquareObj = extSquare

     i = 0
    'Перебираем все блоки и для каждого из них строим зоны горения
    Do
        For Each block In extSquare.extinguishingBlocks
            If block.alreadyCalculated = False Then
                FillExtSquareOneStep block, i
            End If
        Next block
        
        'Фиксируем текущие состояния рассчитанных клеток
        For Each block In extSquare.extinguishingBlocks
            SetCellsToCalculated block.tmpColl
            block.AddToExtCellsCollection block.tmpColl
        Next block
        
        i = i + 1

        If i > 10000 Then
            MsgBox "Аварийный выход в блоке расчета площадей тушения"
        End If
    Loop Until extSquareObj.IsAllBlocksCalculated

    For Each block In extSquare.extinguishingBlocks
        FillActiveCellsSquare block
    Next block
End Sub

'-------------------Отрисовка площади тушения------------------------------
Public Function DrawFireSquare(ByRef extSquare As c_ExtinguishingSquare) As Collection
'Отрисовываем площадь тушения
Dim block As c_ExtinguishingBlock
Dim cell As c_Cell
Dim squareShapesCol As Collection

    Set squareShapesCol = New Collection
    
    'Перебираем все блоки и для каждого из них строим зоны горения
    For Each block In extSquare.extinguishingBlocks
        Set conturPoints = New Collection
        
        If block.extinguishingCells.Count > 0 Then
            For Each cell In block.extinguishingCells
                If block.IsCellInThisPerimeter(cell) Then
                    
                    BuildConturElement cell, block
                End If
            Next cell
    
            'Отрисовываем контуры площади тушения
            squareShapesCol.Add DrawSquare(block)

        End If
    Next block

Set DrawFireSquare = squareShapesCol
End Function

Private Sub BuildConturElement(ByRef cell As c_Cell, ByRef block As c_ExtinguishingBlock)
'Строим контур данной клетки
Dim neighbour(0 To 3) As Boolean
Dim pnt As c_Point
Dim pnt2 As c_Point
    
    'Определяем ортогональных соседей клетки
    neighbour(0) = block.IsCellInThisBlock(cell.NCell(l))   '.IsCellMaxFire
    neighbour(1) = block.IsCellInThisBlock(cell.NCell(u))   'cell.NCell(u).IsCellMaxFire
    neighbour(2) = block.IsCellInThisBlock(cell.NCell(r))   'cell.NCell(r).IsCellMaxFire
    neighbour(3) = block.IsCellInThisBlock(cell.NCell(d))   'cell.NCell(d).IsCellMaxFire
    
    'Проверяем все четыре ортогональных смежных клетки и для каждой из них при необходимости добавляем точку
    '---Проверяем каждую из соседних клеток, на то, горит ли она по максимуму
    If Not neighbour(0) Then
        Set pnt = GetConturPoint(cell.x, cell.y)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x, cell.y + 1)
    End If
    If Not neighbour(1) Then
        Set pnt = GetConturPoint(cell.x, cell.y + 1)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x + 1, cell.y + 1)
    End If
    If Not neighbour(2) Then
        Set pnt = GetConturPoint(cell.x + 1, cell.y + 1)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x + 1, cell.y)
    End If
    If Not neighbour(3) Then
        Set pnt = GetConturPoint(cell.x + 1, cell.y)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x, cell.y)
    End If

End Sub


Private Sub AddPointToCollection(ByRef thisColl As Collection, ByVal pnt As c_Point)
'Добавляем в указанную коллекцию указанную точку
    On Error Resume Next
    thisColl.Add pnt, pnt.x & ":" & pnt.y
End Sub
Private Sub RemovePointFromCollection(ByRef thisColl As Collection, ByVal pnt As c_Point)
'Удаляем указанную точку из указанной коллекции
    On Error Resume Next
    thisColl.Remove pnt.x & ":" & pnt.y
End Sub


Private Function GetConturPoint(ByVal x As Long, ByVal y As Long) As c_Point
Dim pnt As c_Point
    On Error GoTo EX
    Set GetConturPoint = conturPoints.item(x & ":" & y)

Exit Function
EX:
    Set pnt = New c_Point
    pnt.SetData x, y
    AddPointToCollection conturPoints, pnt
    Set GetConturPoint = pnt
End Function



Public Function DrawSquare(ByRef block As c_ExtinguishingBlock) As Visio.Shape
'Отрисовываем площадь используя для этого построение DrawPoiyLine
Dim pnt As c_Point
Dim pntStart As c_Point

Dim shp As Visio.Shape
Dim perimterPnts() As Double
Dim i As Long

    On Error GoTo EX

'#С PerimeterCellsFix
    'Исправляем последовательность точек
    PerimeterCellsFix
'    PerimeterCellsFix2  'Скругление
    
    
    Dim coll As Collection
    For Each coll In conturPoints
        If coll.Count > 1 Then
            i = 0
            ReDim perimterPnts(coll.Count * 2 + 1)
        
            'Инициализируем массив с координатами периметра
            For Each pnt In coll
                perimterPnts(i) = (pnt.x * grain) / mmInInch
                i = i + 1
                perimterPnts(i) = (pnt.y * grain) / mmInInch
                i = i + 1
                
                If i > 100000 Then
                    Debug.Print "Переполнение в стеке отрисовки периметра"    ', vbCritical, "Служебное сообщение"
                    Stop
                    
                    Set DrawSquare = Nothing
                    Exit Function
                End If
            Next pnt
            '---Заключительная точка, замыкающая фигуру
            Set pnt = coll(1)
            perimterPnts(i) = (pnt.x * grain) / mmInInch
            i = i + 1
            perimterPnts(i) = (pnt.y * grain) / mmInInch
            
            'Рисуем фигуру по массиву точек и придаем ей определенные свойства
            Set shp = ActivePage.DrawPolyline(perimterPnts, 0)
                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("ExtSquareTmp")
        End If
    Next coll
    
    Set fireConturs = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "ExtSquareTmp")
    If fireConturs.Count > 0 Then
        fireConturs.Combine
        Set shp = Application.ActiveWindow.Selection(1)
        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Площадь тушения")
        shp.Cells("LineWeight").FormulaU = "0.24 pt"
        shp.Cells("LinePattern").FormulaU = "0"
        shp.Cells("LineColor").FormulaU = "2"
        shp.Cells("FillForegnd").FormulaU = "2"
        shp.Cells("FillForegndTrans").FormulaU = "80%"
        shp.Cells("FillBkgndTrans").FormulaU = "80%"
        
        'Указываем полученной фигуре свойство площади тушения
        shp.AddNamedRow visSectionProp, "ExtSquare", visTagDefault
        shp.CellsSRC(visSectionProp, 0, visCustPropsValue).FormulaU = str(block.GetSquare)
        shp.CellsSRC(visSectionProp, 0, visCustPropsType).FormulaU = 2
        shp.CellsSRC(visSectionProp, 0, visCustPropsLabel).FormulaU = """Площадь тушения"""
    End If

Set DrawSquare = shp
Exit Function
EX:
    Set DrawSquare = Nothing
End Function

Private Sub PerimeterCellsFix()
'Исправляем последовательность точек
Dim pnt As c_Point
Dim tmpPnt As c_Point
Dim curDir As Directions
Dim newPerimeters As Collection
Dim newPerimeterPoints As Collection
Dim i As Long

    Set newPerimeters = New Collection

    Do While conturPoints.Count > 0
        Set newPerimeterPoints = New Collection
        curDir = s
        
        Set pnt = conturPoints(1)
        Set tmpPnt = New c_Point
        tmpPnt.SetData pnt.x, pnt.y
        
        i = 0
        Do
            Set pnt = pnt.GetNextPoint
            'Проверяем в каком направлении сейчас двигается линия контура
            If Not pnt.Dir = curDir Then    'Если направление поменялось - добавляем новую точку в коллекцию и меняем текущее напарвление
                AddPointToCollection newPerimeterPoints, pnt
    
                curDir = pnt.Dir
            End If
            
            RemovePointFromCollection conturPoints, pnt
            
            i = i + 1
            If i > 100000 Then
                Debug.Print "Переполнение в стеке отрисовки периметра"    ', vbCritical, "Служебное сообщение"
                Stop
                Exit Sub
            End If
        
        Loop Until pnt.isEqual(tmpPnt)
        
        newPerimeters.Add newPerimeterPoints
    Loop
    

    'Обновляем коллекцию
    Set conturPoints = newPerimeters
End Sub













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
        Set tmpColl = OneStep(tmpColl, block.fireWay, (i + 1) * grain)
        
        'Указываем для всех клеток полученной коллекции, что они уже рассчитаны
        SetCellsToCalculated tmpColl

        'Добавляем клетки из полученной коллекции в коллекцию клеток площади тушения блока
        block.AddToExtCellsCollection tmpColl
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
    If block.tmpColl.Count = 0 Then
        SetCellsToCalculated block.fireFrontCells
        
        Set block.tmpColl = block.fireFrontCells
    End If
    
    'Рассчитываем площадь тушения (коллекция block.extinguishingCells)
    Set block.tmpColl = OneStep(block.tmpColl, block.fireWay, (i + 1) * grain)
    
    'Если в итоговой коллекции клеток нет ни одной, значит блок рассчитан
    If block.tmpColl.Count = 0 Then
        block.alreadyCalculated = True
    End If
    
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
    
    If calculatedCells(cell.x, cell.y) = 1 And matrix(cell.x, cell.y) > wayLen Then
        calculatedCells(cell.x, cell.y) = 0
        'Удаляем из имеющихся блоков
        extSquareObj.DeleteCellFromEveryBlock cell.x, cell.y
    End If
    
    'Если путь превышает предельное значение или текущее максимальное значение или уже рассчитвана - не учитываем его
    If wayLen >= maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Sub

    'Если больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) >= wayLen Then
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(cell.x, cell.y) = wayLen
        'Удаляем из имеющихся блоков
        extSquareObj.DeleteCellTmpFromEveryBlock cell.x, cell.y

        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If

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



