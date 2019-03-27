VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FrontDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant
Private dimensionByX As Long                                        'Размерность матрицы по ширине (x)
Private dimensionByY As Long                                        'Размерность матрицы по высоте (y)

Public grain As Integer



'---------------------------Класс отвечающий за формирование коллекции блоков горения---------------------------

Private Enum CellType
    notFireFront = 0
    fireFrontNotInBlock = 1
    fireFrontInBlock = 2
End Enum

Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'Активируем матрицу в соответсвии с переданными размерами
    dimensionByX = x
    dimensionByY = y
    ReDim matrix(x, y)
End Sub

Public Sub SetStartState(ByRef frontCellsCollection As Collection)
'Устанавливаем стартовое состояние матрицы
    '1 Задаем стартовое состояние матрицы 5 - путь
    For Each cell In frontCellsCollection
        matrix(cell.x, cell.y) = CellType.fireFrontNotInBlock
    Next cell
End Sub




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection, _
                    ByRef nozzle As c_FireNozzleModel)
'Разбиваем площадь пожара на блоки
Dim block As c_ExtinguishingBlock

    '1 проверяем, не входит ли точка контакта зоны работы ствола в фрагмент фронта пожара, который уже распределен
    Set block = CheckBlockExists(extSquare, nozzle.fireContactCell)
    If block Is Nothing Then
        '2 разбить коллекцию фронта на участки не связанные между собой
        BreakFront extSquare, frontCellsCollection, nozzle
    Else
        '3 Добавляем данные в существующий блок
        block.expenseFact = block.expenseFact + nozzle.waterExpense
    End If
    


End Sub




Private Sub BreakFront(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection, _
                    ByRef nozzle As c_FireNozzleModel)
'Разбиваем фронт пожара на участки
'Dim cell As c_Cell
Dim extBlock As c_ExtinguishingBlock
Dim tmpColl As Collection

    'Перебираем все клетки в коллекции стартовых клеток
'    For Each cell In startCellsCollection
        
        'если клетка еще не распределена по блокам, начинаем от нее рекурсивный поиск всех клеток плоащди тушения данного блока
'        If matrix(startCell.x, startCell.y) = CellType.fireFrontNotInBlock Then
            'создаем новую коллекцию
            Set tmpColl = New Collection
            'Для новой коллекции выполняем рекурсивное наполнение от текущей ячейки
            TryAddCell tmpColl, nozzle.fireContactCell
            
            'Создаем новый блок и передаем ему временную коллекцию клеток фронта
            Set extBlock = New c_ExtinguishingBlock
            Set extBlock.fireFrontCells = tmpColl
            extBlock.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
            extBlock.AddToExtCellsCollection tmpColl
'            AddCellsToCollection extBlock.extinguishingCells, tmpColl
            '---Устанавливаем глубину тушения
                If nozzle.nozzleType = waterHand Then
                    extBlock.fireWay = 5000     'Для ручных стволов
                Else
                    extBlock.fireWay = 10000     'Для лафетных стволов
                End If
                extBlock.expenseFact = extBlock.expenseFact + nozzle.waterExpense
                extBlock.grain = grain
            
            'Добавляем созданный блок в коллекцию блоков модели площади тушения
            extSquare.AddExtinguishingBlock extBlock
'        End If
'    Next cell
        
End Sub

Private Sub TryAddCell(ByRef thisColl As Collection, ByVal cell As c_Cell)
'Пытаемся добавить клетку в коллекцию (рекурсия)
    
    'Если указанная точка находится за пределами матрицы - выходим из процедуры
    If cell.x < 0 Or cell.y < 0 Or cell.x > dimensionByX - 1 Or cell.y > dimensionByY - 1 Then Exit Sub
    
    'Проверяем является ли данная клетка клеткой фронта пожара
    If matrix(cell.x, cell.y) = CellType.notFireFront Then Exit Sub
    
    'Если является проверяем, нет ли ее уже в коллекции
    If matrix(cell.x, cell.y) = CellType.fireFrontInBlock Then Exit Sub
    
    'Добавляем клетку в коллецию
    AddCellToCollection thisColl, cell
    matrix(cell.x, cell.y) = CellType.fireFrontInBlock
    
    'Если нет - рекурсия на следующие клетки
        TryAddCell thisColl, NextCell(cell, l)
        TryAddCell thisColl, NextCell(cell, lu)
        TryAddCell thisColl, NextCell(cell, u)
        TryAddCell thisColl, NextCell(cell, ru)
        TryAddCell thisColl, NextCell(cell, r)
        TryAddCell thisColl, NextCell(cell, rd)
        TryAddCell thisColl, NextCell(cell, d)
        TryAddCell thisColl, NextCell(cell, ld)
   
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

Private Function CheckBlockExists(ByRef extSquare As c_ExtinguishingSquare, ByRef cell As c_Cell) As c_ExtinguishingBlock
'Пытаемся найти блок к которому может относиться пришедшая клетка. Еслитн аходим возвращаем ссылку на него, иначе - Nothing
Dim blck As c_ExtinguishingBlock

    For Each blck In extSquare.extinguishingBlocks
        If blck.IsCellInThisFront(cell) Then
            Set CheckBlockExists = blck
            Exit Function
        End If
    Next blck
Set CheckBlockExists = Nothing
End Function

'Public Function DrawBlockExists(ByRef extSquare As c_ExtinguishingSquare, ByRef cell As c_Cell) As c_ExtinguishingBlock
''Пытаемся найти блок к которому может относиться пришедшая клетка. Еслитн аходим возвращаем ссылку на него, иначе - Nothing
'Dim blck As c_ExtinguishingBlock
'
'Dim c As c_Cell
'Dim shp As Visio.Shape
'Const mmInInch As Double = 25.4
'
'    For Each blck In extSquare.extinguishingBlocks
'        For Each c In blck.fireFrontCells
'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle(c.x * grain / mmInInch, c.y * grain / mmInInch, _
'                    (c.x + 1) * grain / mmInInch, (c.y + 1) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Front")
'            shp.Cells("LinePattern").FormulaU = "0"
'            shp.Cells("FillForegnd").FormulaU = "3"
'        Next c
'
'
''        If blck.IsCellInThisFront(cell) Then
''            Set CheckBlockExists = blck
''            Exit Function
''        End If
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

Public Property Get GetMatrix() As Variant
    GetMatrix = matrix
End Property
