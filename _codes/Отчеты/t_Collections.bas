Attribute VB_Name = "t_Collections"
'--------------------------------Коллекции-------------------------------------
Public Sub AddCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Add newCollection items to oldCollection
Dim item As Object

    For Each item In newCollection
        oldCollection.Add item
    Next item
End Sub

Public Sub AddUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Add newCollection items (with unique .ID prop) to oldCollection
Dim item As Object

    For Each item In newCollection
       AddUniqueCollectionItem oldCollection, item
    Next item
End Sub

Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As Object)
'Add item (with unique .key prop) to oldCollection
    On Error GoTo EX
    
    oldCollection.Add item, CStr(item.ID)

Exit Sub
EX:
'    Debug.Print "Item with key='" & item.ID & "' is already exists!)"
End Sub

Public Sub SetCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Refresh old collection items with items from newCollection
Dim item As Object
    On Error GoTo EX
    
    Set oldCollection = New Collection

    For Each item In newCollection
        oldCollection.Add item, item.ID
    Next item
    
Exit Sub
EX:
'    Debug.Print "Item with key='" & item.ID & "' is already exists!)"
End Sub

Public Sub SetUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Refresh old collection items with items (with unique .key prop) from newCollection
Dim item As Object

    Set oldCollection = New Collection

    For Each item In newCollection
        AddUniqueCollectionItem oldCollection, item
    Next item

End Sub

Public Sub RemoveFromCollection(ByRef oldCollection As Collection, ByRef item As Object)
'Remove specific item (with unique .ID prop) from collection
    On Error Resume Next
    oldCollection.Remove CStr(item.ID)
End Sub

Public Function GetFromCollection(ByRef coll As Collection, ByVal ID As String) As Object
'Get specific item (with unique .ID prop) from collection
Dim item As Object

    On Error GoTo EX
    Set item = coll.item(ID)
    If Not item Is Nothing Then
        Set GetFromCollection = item
    Else
        Set GetFromCollection = Nothing
    End If
    
Exit Function
EX:
    Set GetFromCollection = Nothing
End Function

Public Function IsInCollection(ByRef coll As Collection, obj As Object) As Boolean
'Check item (with unique .ID prop) existance in collection
Dim item As Object

    On Error GoTo EX
    
    Set item = coll.item(CStr(obj.ID))
    If Not item Is Nothing Then
        IsInCollection = True
    Else
        IsInCollection = False
    End If
    
Exit Function
EX:
    IsInCollection = False
End Function


Public Function FilterShapes(ByRef shpColl As Variant, ByVal filterStr As String, _
            Optional ByVal d_elem As String = ";", Optional ByVal d_val As String = ":") As Collection
'Возвращает коллекцию фигур отфильтрованных по условию в котором указываются ячейки и их значения
'shpColl - коллекция фигур которую следует отфильтровать
'filterStr - строка фильтра
'd_elem разделитель пар ячейка/значение
'd_val разделитель ячейки и значения в паре ячейка/значение - если после разделителя нет ничего, то проверяется просто наличие такой ячейки
'Применение: GetGFSShapes("User.IndexPers:500;User.IndexPers:1", ":", ";")
'На вход принимаются коллекции только Visio.Shape. Все прочие объкуты игнорируются!
Dim shp As Visio.Shape
Dim filters() As String
Dim filterItem() As String
Dim filterItemCellName As String
Dim filterItemCellValue As String
Dim i As Integer
Dim tmpColl As Collection
    
    On Error GoTo EX
    
    Set tmpColl = New Collection
    
    filters = Split(filterStr, d_elem)
    
'    Debug.Print TypeName(shpColl)
    For Each shp In shpColl
        'Проверка на shp!
        If TypeName(shp) = "Shape" Then
            For i = 0 To UBound(filters)
                filterItem = Split(filters(i), d_val)
                filterItemCellName = filterItem(0)
                filterItemCellValue = filterItem(1)
                
                If filterItemCellValue = "" Then
                    If ShapeHaveCell(shp, filterItemCellName) Then AddUniqueCollectionItem tmpColl, shp
                Else
                    If ShapeHaveCell(shp, filterItemCellName, filterItemCellValue) Then AddUniqueCollectionItem tmpColl, shp
                End If
            Next i
        End If
    Next shp
    
    Set FilterShapes = tmpColl
    
Exit Function
EX:
    Set FilterShapes = New Collection
End Function

Public Function FilterShapesAnd(ByRef shpColl As Variant, ByVal filterStr As String, _
            Optional ByVal d_elem As String = ";", Optional ByVal d_val As String = ":") As Collection
'Возвращает коллекцию фигур отфильтрованных по условию в котором указываются ячейки и их значения
'shpColl - коллекция фигур которую следует отфильтровать
'filterStr - строка фильтра
'd_elem разделитель пар ячейка/значение
'd_val разделитель ячейки и значения в паре ячейка/значение - если после разделителя нет ничего, то проверяется просто наличие такой ячейки
'Применение: GetGFSShapes("User.IndexPers:500;User.IndexPers:1", ":", ";")
'На вход принимаются коллекции только Visio.Shape. Все прочие объкуты игнорируются!
Dim shp As Visio.Shape
Dim filters() As String
Dim filterItem() As String
Dim filterItemCellName As String
Dim filterItemCellValue As String
Dim i As Integer
Dim tmpColl As Collection
Dim approved As Boolean
    
    On Error GoTo EX
    
    Set tmpColl = New Collection
    
    filters = Split(filterStr, d_elem)
    
'    Debug.Print TypeName(shpColl)
    For Each shp In shpColl
        'Проверка на shp!
        If TypeName(shp) = "Shape" Then
            approved = True
            For i = 0 To UBound(filters)
                filterItem = Split(filters(i), d_val)
                filterItemCellName = filterItem(0)
                filterItemCellValue = filterItem(1)
                
                If filterItemCellValue = "" Then
                    If Not ShapeHaveCell(shp, filterItemCellName) Then
                        approved = False
                        Exit For
                    End If
                Else
                    If Not ShapeHaveCell(shp, filterItemCellName, filterItemCellValue) Then
                        approved = False
                        Exit For
                    End If
                End If
            Next i
            If approved Then
                AddUniqueCollectionItem tmpColl, shp
            End If
        End If
    Next shp
    
    Set FilterShapesAnd = tmpColl
    
Exit Function
EX:
    Set FilterShapesAnd = New Collection
End Function

'------------------------Агрегатные функции по коллекциям-----------------------------------
'Public Function CellSum(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VbVarType = vbSingle) As Variant
Public Function CellSum(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'Возвращает сумму значений ячеек с именем cellName всех фигур в коллекции shpColl
Dim shp As Visio.Shape
Dim tmpval As Variant

    For Each shp In shpColl
        'Проверка на shp!
        If TypeName(shp) = "Shape" Then
            tmpval = tmpval + cellVal(shp, cellName, returnType)
        End If
    Next shp
CellSum = tmpval
End Function

Public Function CellMax(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'Возвращает максимальное значение ячеек с именем cellName всех фигур в коллекции shpColl
Dim shp As Visio.Shape
Dim tmpval As Variant
Dim cVal As Variant
    
    tmpval = 0
    For Each shp In shpColl
        'Проверка на shp!
        If TypeName(shp) = "Shape" Then
            cVal = cellVal(shp, cellName, returnType)
            If tmpval < cVal Then
                tmpval = cVal
            End If
        End If
    Next shp
CellMax = tmpval
End Function

Public Function CellMin(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'Возвращает максимальное значение ячеек с именем cellName всех фигур в коллекции shpColl
Dim shp As Visio.Shape
Dim tmpval As Variant
Dim cVal As Variant
Dim start As Boolean
    
    start = True
    For Each shp In shpColl
        'Проверяем на первый запуск
        If start Then
            tmpval = cellVal(shp, cellName, returnType)
            start = Not start
        Else
            'Проверка на shp!
            If TypeName(shp) = "Shape" Then
                cVal = cellVal(shp, cellName, returnType)
                If tmpval > cVal Then
                    tmpval = cVal
                End If
            End If
        End If
    Next shp
CellMin = tmpval
End Function

Public Function CellAvg(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'Возвращает максимальное значение ячеек с именем cellName всех фигур в коллекции shpColl
Dim shp As Visio.Shape
Dim total As Variant
Dim i As Long
    
    i = 0
    For Each shp In shpColl
        'Проверка на shp!
        If TypeName(shp) = "Shape" Then
            total = total + cellVal(shp, cellName, returnType)
        End If
        
        i = i + 1
    Next shp
CellAvg = total / i
End Function

'Public Sub TTT()
'Dim c As Collection
'
'Set c = A.Refresh(1).GetGFSShapes("User.IndexPers:" & indexPers.ipStvolRuch & ";User.IndexPers:" & indexPers.ipStvolLafVoda)
'
'Debug.Print CellAvg(c, "User.PodOut", visNumber)
'
'End Sub

'Public Sub TTT()
'Dim c As Collection
'Dim s As Collection
'Dim shp As Visio.Shape
'
'    Set c = New Collection
'    For Each shp In Application.ActivePage.Shapes
'        c.Add shp
'    Next shp
'
'    Debug.Print FilterShapes(Application.ActivePage.Shapes, "User.IndexPers:62").count
''    Debug.Print FilterShapes(c, "User.IndexPers:500;User.IndexPers:1").count
''    Debug.Print FilterShapes(c, "User.IndexPers:").count
'End Sub
