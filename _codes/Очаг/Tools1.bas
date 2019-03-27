Attribute VB_Name = "Tools1"


'--------------------------------Работа с точками-------------------------------------
'Public Function SpecialRound(ByVal originalVal As Double, ByVal koeff As Single) As Integer
''Округляем со указанным коеффициентом смещения порогового значения
'Dim bottomVal As Long
'Dim topVal As Long
'Dim midVal As Double
'
'    bottomVal = Int(originalVal)
'    topVal = bottomVal + 1
'    midVal = bottomVal + koeff
'
'    If originalVal < midVal Then
'        SpecialRound = bottomVal
'    Else
'        SpecialRound = topVal
'    End If
'
'End Function

'-------------------------------Работа с массивами---------------------------------------------------
Public Function SplitToDouble(ByVal str As String, ByVal delimiter As String) As Double()
'Возвращает строку разбитую в массив Double
Dim tempArr() As String
Dim tempArrD() As Double
Dim i As Long

    tempArr = Split(str, delimiter)
    '---Получаем массив Double
    ReDim tempArrD(UBound(tempArr))
    For i = 0 To UBound(tempArr)
        tempArrD(i) = CDbl(tempArr(i))
    Next i
    
    SplitToDouble = tempArrD
    
End Function

Public Function MakeArrayForDraw(ByRef arrD() As Double) As Double()
'Возвращаем готовый для отрисовки массив
    ReDim Preserve arrD(UBound(arrD) + 2)
    
    arrD(UBound(arrD) - 1) = arrD(0)
    arrD(UBound(arrD)) = arrD(1)
    
    MakeArrayForDraw = arrD
End Function

'Public Function GetArrayForDraw(ByVal str As String, ByVal delimiter As String) As Double()
''Возвращаем готовый для отрисовки массив
'    SplitToDouble str, delimiter
'
'End Function


'--------------------------------Коллекции-------------------------------------
Public Sub AddCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Добавляем элементы новой коллекции в старую
Dim GenPointItem As c_Point

    '---Перебираем все элементы в новой коллекции и добавляем их в старую
    For Each GenPointItem In newCollection
        oldCollection.Add GenPointItem
    Next GenPointItem
End Sub

Public Sub AddUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Добавляем новые элементы новой коллекции в старую
Dim GenPointItem As c_Point

    '---Перебираем все элементы в новой коллекции и добавляем их в старую
    For Each GenPointItem In newCollection
       AddUniqueCollectionItem oldCollection, GenPointItem
    Next GenPointItem
End Sub

Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As c_Point)
Dim GenPointItem As c_Point

    '---Перебираем все элементы в новой коллекции и добавляем их в старую
    For Each GenPointItem In oldCollection
        If GenPointItem.x = item.x And GenPointItem.y = item.y Then Exit Sub
    Next GenPointItem

    oldCollection.Add item
End Sub

Public Sub SetCollection(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Обновляем старую коллекцию в соответствии со значениями новой коллекции
Dim item As Variant

    Set oldCollection = New Collection

    '---Перебираем все элементы в новой коллекции и добавляем их в старую
    For Each item In newCollection
        oldCollection.Add item
    Next item

    '---очищаем новую коллекцию
End Sub

Public Sub RemoveFromCollection(ByRef coll As Collection, element As Variant)
Dim item As Variant
Dim i As Integer

    i = 0

    For Each item In coll
        If item.x = element.x And item.y = element.y Then
            coll.Remove i + 1
            Exit Sub
        End If
        i = i + 1
    Next item
End Sub

Public Function IsInCollection(ByRef coll As Collection, point As c_Point) As Boolean
Dim item As c_Point

    For Each item In coll
        If item.isEqual(point) Then
            IsInCollection = True
            Exit Function
        End If
    Next item

IsInCollection = False
End Function

Public Sub NormalizeCollection(ByRef col As Collection)
'Нормализуем коллекцию координат точек для отрисовки
Dim pnt1 As c_Point
Dim pnt2 As c_Point
Dim pnt3 As c_Point
Dim vector1 As c_Point
Dim vector2 As c_Point
Dim i As Long
Dim angle1 As Double
Dim angle2 As Double
    
    i = 1
    Do While i < col.Count - 1
        Set pnt1 = col.item(i)
        Set pnt2 = col.item(i + 1)
        Set pnt3 = col.item(i + 2)
        Set vector1 = New c_Point
            vector1.SetData (pnt2.x - pnt1.x), (pnt2.y - pnt1.y)
        Set vector2 = New c_Point
            vector2.SetData (pnt3.x - pnt2.x), (pnt3.y - pnt2.y)
        
        If vector1.GetTan = vector2.GetTan Then
            col.Remove i + 1
        Else
            i = i + 1
        End If
    Loop
    
End Sub

'--------------------------------Работа со слоями-------------------------------------
Public Function GetLayerNumber(ByRef layerName As String) As Integer
Dim layer As Visio.layer

    For Each layer In Application.ActivePage.Layers
        If layer.Name = layerName Then
            GetLayerNumber = layer.Index - 1
            Exit Function
        End If
    Next layer
    
    Set layer = Application.ActivePage.Layers.Add(layerName)
    GetLayerNumber = layer.Index - 1
End Function

'---------------------------------------Служебные функции и проки--------------------------------------------------
Public Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'Функция возвращает угол относительно родительского элемента
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

'Set Shape = Nothing
End Function

Public Sub ClearLayer(ByVal layerName As String)
'Удаляем фигуры указанного слоя
    On Error Resume Next
    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, layerName)
    vsoSelection.Delete
End Sub

Public Function ShapeIsLine(ByRef shp As Visio.Shape) As Boolean
'Функция возвращает истина, если переданная фигура - простая прямая линия, Ложь - если иначе
Dim isLine As Boolean
Dim isStrait As Boolean
    
    ShapeIsLine = False
    
    On Error GoTo EX
    
    If shp.RowCount(visSectionFirstComponent) <> 3 Then Exit Function       'Строк в секции геометрии больше или меньше двух
    If shp.RowType(visSectionFirstComponent, 2) <> 139 Then Exit Function   '139 - LineTo
    
ShapeIsLine = True
Exit Function

EX:
ShapeIsLine = False
End Function

Public Function GetDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
Dim catet1 As Double
Dim catet2 As Double
    
    catet1 = x2 - x1
    catet2 = y2 - y1
    
    GetDistance = Sqr(catet1 ^ 2 + catet2 ^ 2)
End Function

Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function

Public Function PFB_isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - дверной проем, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isDoor = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ДВЕРЬ или ПРОЕМ
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Or aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25) Then
        PFB_isDoor = True
        Exit Function
    End If
PFB_isDoor = False
End Function

Public Function PFI_FirstSectionCount(ByRef aO_Shape As Visio.Shape) As Integer
'Функция возвращает количество графических секций
Dim i As Integer

    i = 0
    Do While aO_Shape.SectionExists(visSectionFirstComponent + i, 0)
        i = i + 1
    Loop
    
PFI_FirstSectionCount = i
End Function


