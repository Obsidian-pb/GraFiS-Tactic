VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PerimeterDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conturPoints As Collection                                  'Коллекция точек периметра

Public grain As Integer
Const mmInInch As Double = 25.4

'--------------------------------Демон построитель периметра (площади пожара)-------------------------------------------



Private Sub Class_Initialize()
    Set conturPoints = New Collection
    Set perimeterCells = New Collection
End Sub


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
    BuildConturLine perimeterCellsCollection
    
    
    
    

End Sub



Private Sub BuildConturLine(ByRef perimeterCellsCollection As Collection)
'Формируем контур фигуры периметра
Dim cell As c_Cell
    
    For Each cell In perimeterCellsCollection
        BuildConturElement cell
    Next cell
End Sub

Private Sub BuildConturElement(ByRef cell As c_Cell)

Dim neighbour(0 To 3) As Boolean
Dim pnt As c_Point
Dim pnt2 As c_Point
    
    'Определяем ортогональных соседей клетки
    neighbour(0) = cell.NCell(l).IsCellMaxFire
    neighbour(1) = cell.NCell(u).IsCellMaxFire
    neighbour(2) = cell.NCell(r).IsCellMaxFire
    neighbour(3) = cell.NCell(d).IsCellMaxFire
    
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




'-----------Собственно отрисовка площади------------
Public Function DrawSquare() As Visio.Shape
'Отрисовываем площадь используя для этого построение DrawPoiyLine
Dim pnt As c_Point
Dim pntStart As c_Point

Dim shp As Visio.Shape
Dim perimterPnts() As Double
'Dim i As Long


''#Без PerimeterCellsFix
'    i = 0
'    ReDim perimterPnts(conturPoints.Count * 2 + 1)
'
'    Set pnt = conturPoints(1)
'    Set pntStart = New c_Point
'    pntStart.SetData pnt.x, pnt.y
'
'    perimterPnts(i) = (pnt.x * grain) / mmInInch
'    i = i + 1
'    perimterPnts(i) = (pnt.y * grain) / mmInInch
'    i = i + 1
'
'    Do
'        Set pnt = pnt.GetNextPoint
'        perimterPnts(i) = (pnt.x * grain) / mmInInch
'        i = i + 1
'        perimterPnts(i) = (pnt.y * grain) / mmInInch
'        i = i + 1
'
'        If i > 10000 Then
'            MsgBox "Переполнение в стеке отрисовки периметра", vbCritical, "Служебное сообщение"
'            Exit Do
'        End If
'    Loop Until pnt.isEqual(pntStart)

'#С PerimeterCellsFix
    'Исправляем последовательность точек
    PerimeterCellsFix

    i = 0
    ReDim perimterPnts(conturPoints.Count * 2 + 1)

    'Инициализируем массив с координатами периметра
    For Each pnt In conturPoints
        perimterPnts(i) = (pnt.x * grain) / mmInInch
        i = i + 1
        perimterPnts(i) = (pnt.y * grain) / mmInInch
        i = i + 1
    Next pnt
    '---Заключительная точка, замыкающая фигуру
    Set pnt = conturPoints(1)
    perimterPnts(i) = (pnt.x * grain) / mmInInch
    i = i + 1
    perimterPnts(i) = (pnt.y * grain) / mmInInch
    
    

    'Рисуем фигуру по массиву точек и придаем ей определенные свойства
    Set shp = ActivePage.DrawPolyline(perimterPnts, 0)
        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
        shp.Cells("LinePattern").FormulaU = "0"
        shp.Cells("FillForegnd").FormulaU = "2"
        shp.Cells("FillForegndTrans").FormulaU = "50%"
        shp.Cells("FillBkgndTrans").FormulaU = "50%"


Set DrawSquare = shp
End Function

Private Sub PerimeterCellsFix()
'Исправляем последовательность точек
Dim pnt As c_Point
Dim tmpPnt As c_Point
Dim curDir As Directions
Dim newPerimeterPoints As Collection

    Set newPerimeterPoints = New Collection
    curDir = s

'    For Each pnt In conturPoints
'        If Not pnt.Dir = curDir Then
'            AddPointToCollection newPerimeterPoints, pnt
'
'            curDir = pnt.Dir
'
'        End If
'    Next pnt
    
    Set pnt = conturPoints(1)
    Set tmpPnt = New c_Point
    tmpPnt.SetData pnt.x, pnt.y
    
    Do
        Set pnt = pnt.GetNextPoint
        'Проверяем в како направлении сейчас двигается линия контура
        If Not pnt.Dir = curDir Then    'Если направление поменялось - добавляем новую точку в коллекцию и меняем текущее напарвление
            AddPointToCollection newPerimeterPoints, pnt

            curDir = pnt.Dir
        End If

'        If i > 10000 Then
'            MsgBox "Переполнение в стеке отрисовки периметра", vbCritical, "Служебное сообщение"
'            Exit Do
'        End If
    Loop Until pnt.isEqual(tmpPnt)
    
    

    'Обновляем коллекцию
    Set conturPoints = newPerimeterPoints
End Sub



