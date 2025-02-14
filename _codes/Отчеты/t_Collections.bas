Attribute VB_Name = "t_Collections"
'--------------------------------���������-------------------------------------
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

'Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As Object, Optional ByVal key As String = "")
Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByVal item As Variant, Optional ByVal key As String = "")
'Add item (with unique .key prop) to oldCollection
'Dim i As Integer
'Dim s As String
    On Error GoTo ex
    
    If key = "" Then
        oldCollection.Add item, CStr(item.ID)
    Else
'        s = key & " "
'        For i = 1 To Len(key)
'            s = s & Asc(Mid(key, i, 1))
'        Next i
'        Debug.Print s
        oldCollection.Add item, CStr(key)
    End If

Exit Sub
ex:
'    Debug.Print "Item with key='" & item.ID & "' is already exists!)"
End Sub

Public Sub SetCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'Refresh old collection items with items from newCollection
Dim item As Object
    On Error GoTo ex
    
    Set oldCollection = New Collection

    For Each item In newCollection
        oldCollection.Add item, item.ID
    Next item
    
Exit Sub
ex:
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

    On Error GoTo ex
    Set item = coll.item(ID)
    If Not item Is Nothing Then
        Set GetFromCollection = item
    Else
        Set GetFromCollection = Nothing
    End If
    
Exit Function
ex:
    Set GetFromCollection = Nothing
End Function

Public Function IsInCollection(ByRef coll As Collection, obj As Object) As Boolean
'Check item (with unique .ID prop) existance in collection
Dim item As Object

    On Error GoTo ex
    
    Set item = coll.item(CStr(obj.ID))
    If Not item Is Nothing Then
        IsInCollection = True
    Else
        IsInCollection = False
    End If
    
Exit Function
ex:
    IsInCollection = False
End Function

Public Function IsKeyInCollection(ByRef coll As Collection, key As String) As Boolean
'Check item's key existance in collection
Dim item As Object

    On Error GoTo ex
    
    Set item = coll.item(key)
    If Not item Is Nothing Then
        IsKeyInCollection = True
    Else
        IsKeyInCollection = False
    End If
    
Exit Function
ex:
    IsKeyInCollection = False
End Function

Public Function FilterShapes(ByRef shpColl As Variant, ByVal filterStr As String, _
            Optional ByVal d_elem As String = ";", Optional ByVal d_val As String = ":") As Collection
'���������� ��������� ����� ��������������� �� ������� � ������� ����������� ������ � �� ��������
'shpColl - ��������� ����� ������� ������� �������������
'filterStr - ������ �������
'd_elem ����������� ��� ������/��������
'd_val ����������� ������ � �������� � ���� ������/�������� - ���� ����� ����������� ��� ������, �� ����������� ������ ������� ����� ������
'����������: FilterShapes(A.GFSShapes, "User.IndexPers:500;User.IndexPers:1", ":", ";")
'�� ���� ����������� ��������� ������ Visio.Shape. ��� ������ ������� ������������!
Dim shp As Visio.Shape
Dim filters() As String
Dim filterItem() As String
Dim filterItemCellName As String
Dim filterItemCellValue As String
Dim i As Integer
Dim tmpColl As Collection
    
    On Error GoTo ex
    
    Set tmpColl = New Collection
    
    filters = Split(filterStr, d_elem)
    
'    Debug.Print TypeName(shpColl)
    For Each shp In shpColl
        '�������� �� shp!
        If TypeName(shp) = "Shape" Then
            For i = 0 To UBound(filters)
                filterItem = Split(filters(i), d_val)
                If UBound(filterItem) = 1 Then
                    filterItemCellName = filterItem(0)
                    filterItemCellValue = filterItem(1)
                ElseIf UBound(filterItem) = 0 Then
                    filterItemCellName = filterItem(0)
                    filterItemCellValue = ""
                End If
                
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
ex:
    Set FilterShapes = New Collection
End Function

Public Function FilterShapesAnd(ByRef shpColl As Variant, ByVal filterStr As String, _
            Optional ByVal d_elem As String = ";", Optional ByVal d_val As String = ":") As Collection
'���������� ��������� ����� ��������������� �� ������� � ������� ����������� ������ � �� ��������
'shpColl - ��������� ����� ������� ������� �������������
'filterStr - ������ �������
'd_elem ����������� ��� ������/��������
'd_val ����������� ������ � �������� � ���� ������/�������� - ���� ����� ����������� ��� ������, �� ����������� ������ ������� ����� ������
'����������: GetGFSShapes("User.IndexPers:500;User.IndexPers:1", ":", ";")
'�� ���� ����������� ��������� ������ Visio.Shape. ��� ������ ������� ������������!
Dim shp As Visio.Shape
Dim filters() As String
Dim filterItem() As String
Dim filterItemCellName As String
Dim filterItemCellValue As String
Dim i As Integer
Dim tmpColl As Collection
Dim approved As Boolean
    
    On Error GoTo ex
    
    Set tmpColl = New Collection
    
    filters = Split(filterStr, d_elem)
    
'    Debug.Print TypeName(shpColl)
    For Each shp In shpColl
        '�������� �� shp!
        If TypeName(shp) = "Shape" Then
            approved = True
            For i = 0 To UBound(filters)
                filterItem = Split(filters(i), d_val)
                If UBound(filterItem) = 1 Then
                    filterItemCellName = filterItem(0)
                    filterItemCellValue = filterItem(1)
                ElseIf UBound(filterItem) = 0 Then
                    filterItemCellName = filterItem(0)
                    filterItemCellValue = ""
                End If
                
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
ex:
    Set FilterShapesAnd = New Collection
End Function

Public Function SortCol(ByVal shps As Collection, ByVal sortCellName As String, Optional ByVal desc As Boolean = True, _
                        Optional ByVal dataType As VisUnitCodes = visNumber) As Collection
'������� ���������� ��������������� ��������� �����. ������ ����������� �� �������� � ������ sortCellName - ��� ������, ��� ����
'desc: True - �� �������� � ��������; False - �� �������� � ��������
Dim i As Integer
Dim tmpshp As Visio.Shape
Dim innerShps As Collection     '��������� �� �������� ����� ��������� shps, ��� �������� �� ���������
Dim tmpColl As Collection

    
    Set tmpColl = New Collection
    
    '������� ����� �������� ���������
    Set innerShps = New Collection
    AddUniqueCollectionItems innerShps, shps
    
    Do While innerShps.count > 1
        
        If desc Then
            Set tmpshp = GetMaxShp(innerShps, sortCellName, dataType)
        Else
            Set tmpshp = GetMinShp(innerShps, sortCellName, dataType)
        End If
        
        AddUniqueCollectionItem tmpColl, tmpshp
        RemoveFromCollection innerShps, tmpshp
        
        i = i + 1
        If i > 10000 Then Exit Do
    Loop
    
    Set tmpshp = innerShps(1)
    AddUniqueCollectionItem tmpColl, tmpshp
'    Debug.Print tmpshp.ID & " " & cellVal(tmpshp, sortCellName)
    
    Set SortCol = tmpColl
End Function

Public Function GetMaxShp(ByRef col As Collection, ByVal sortCellName As String, Optional ByVal dataType As VisUnitCodes = visNumber) As Visio.Shape
'���������� ������ � ������������ ��������� � ������ sortCellName
Dim i As Integer
Dim j As Integer
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim shp1Val As Variant
Dim shp2Val As Variant
Const d = ";"
Dim sortCellNames() As String
Dim sortCellNameOne As String
    
    Set shp1 = col(1)
            
            '---��� ������, ���� � �������� ����� ���������� ������� ��������� ��������� ����� (�� ������������!)
            If InStr(1, sortCellName, d) > 0 Then
                sortCellNames = Split(sortCellName, d)
                For j = 0 To UBound(sortCellNames)
                    If ShapeHaveCell(shp1, sortCellNames(j)) Then
                        sortCellNameOne = sortCellNames(j)
                        Exit For
                    End If
                Next j
            Else
                sortCellNameOne = sortCellName
            End If
    
    shp1Val = cellVal(shp1, sortCellNameOne, dataType)
    For i = 1 To col.count
        Set shp2 = col(i)
        
            '---��� ������, ���� � �������� ����� ���������� ������� ��������� ��������� ����� (�� ������������!)
            If InStr(1, sortCellName, d) > 0 Then
'                sortCellNames = Split(sortCellName, d)
                For j = 0 To UBound(sortCellNames)
                    If ShapeHaveCell(shp2, sortCellNames(j)) Then
                        sortCellNameOne = sortCellNames(j)
                        Exit For
                    End If
                Next j
            Else
                sortCellNameOne = sortCellName
            End If
        
        
        shp2Val = cellVal(shp2, sortCellNameOne, dataType)
        If shp2Val > shp1Val Then
            Set shp1 = shp2
            shp1Val = shp2Val
        End If
    Next i
'    Debug.Print shp1.ID & " " & shp1Val
    
Set GetMaxShp = shp1
End Function

Public Function GetMinShp(ByRef col As Collection, ByVal sortCellName As String, Optional ByVal dataType As VisUnitCodes = visNumber) As Visio.Shape
'���������� ������ � ����������� ��������� � ������ sortCellName
Dim i As Integer
Dim j As Integer
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape
Dim shp1Val As Variant
Dim shp2Val As Variant
Const d = ";"
Dim sortCellNames() As String
Dim sortCellNameOne As String
    
    Set shp1 = col(1)
    
            '---��� ������, ���� � �������� ����� ���������� ������� ��������� ��������� ����� (�� ������������!)
            If InStr(1, sortCellName, d) > 0 Then
                sortCellNames = Split(sortCellName, d)
                For j = 0 To UBound(sortCellNames)
                    If ShapeHaveCell(shp1, sortCellNames(j)) Then
                        sortCellNameOne = sortCellNames(j)
                        Exit For
                    End If
                Next j
            Else
                sortCellNameOne = sortCellName
            End If
    
    
    shp1Val = cellVal(shp1, sortCellNameOne, dataType)
'    Debug.Print sortCellNameOne & ", " & shp1Val
    For i = 1 To col.count
        Set shp2 = col(i)
        
            '---��� ������, ���� � �������� ����� ���������� ������� ��������� ��������� ����� (�� ������������!)
            If InStr(1, sortCellName, d) > 0 Then
'                sortCellNames = Split(sortCellName, d)
                For j = 0 To UBound(sortCellNames)
                    If ShapeHaveCell(shp2, sortCellNames(j)) Then
                        sortCellNameOne = sortCellNames(j)
                        Exit For
                    End If
                Next j
            Else
                sortCellNameOne = sortCellName
            End If
        
        
        shp2Val = cellVal(shp2, sortCellNameOne, dataType)
        If shp2Val < shp1Val Then
            Set shp1 = shp2
            shp1Val = shp2Val
        End If
    Next i
'    Debug.Print shp1.ID & " " & shp1Val
    
Set GetMinShp = shp1
End Function

'------------------------���������� ������� �� ����������-----------------------------------
'Public Function CellSum(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VbVarType = vbSingle) As Variant
Public Function CellSum(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'���������� ����� �������� ����� � ������ cellName ���� ����� � ��������� shpColl
Dim shp As Visio.Shape
Dim tmpVal As Variant

    For Each shp In shpColl
        '�������� �� shp!
        If TypeName(shp) = "Shape" Then
            tmpVal = tmpVal + cellVal(shp, cellName, returnType)
        End If
    Next shp
CellSum = tmpVal
End Function

Public Function CellMax(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'���������� ������������ �������� ����� � ������ cellName ���� ����� � ��������� shpColl
Dim shp As Visio.Shape
Dim tmpVal As Variant
Dim cVal As Variant
    
    tmpVal = 0
    For Each shp In shpColl
        '�������� �� shp!
        If TypeName(shp) = "Shape" Then
            cVal = cellVal(shp, cellName, returnType)
            If tmpVal < cVal Then
                tmpVal = cVal
            End If
        End If
    Next shp
CellMax = tmpVal
End Function

Public Function CellMin(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'���������� ������������ �������� ����� � ������ cellName ���� ����� � ��������� shpColl
Dim shp As Visio.Shape
Dim tmpVal As Variant
Dim cVal As Variant
Dim start As Boolean
    
    start = True
    For Each shp In shpColl
        '��������� �� ������ ������
        If start Then
            tmpVal = cellVal(shp, cellName, returnType)
            start = Not start
        Else
            '�������� �� shp!
            If TypeName(shp) = "Shape" Then
                cVal = cellVal(shp, cellName, returnType)
                If tmpVal > cVal Then
                    tmpVal = cVal
                End If
            End If
        End If
    Next shp
CellMin = tmpVal
End Function

Public Function CellAvg(ByRef shpColl As Variant, ByVal cellName As String, Optional ByVal returnType As VisUnitCodes = visNumber) As Variant
'���������� ������������ �������� ����� � ������ cellName ���� ����� � ��������� shpColl
Dim shp As Visio.Shape
Dim total As Variant
Dim i As Long
    
    i = 0
    For Each shp In shpColl
        '�������� �� shp!
        If TypeName(shp) = "Shape" Then
            total = total + cellVal(shp, cellName, returnType)
        End If
        
        i = i + 1
    Next shp
CellAvg = total / i
End Function

Public Function GetUniqueVals(ByRef shpColl As Variant, ByVal cellName As String, _
                              Optional ByVal returnType As VisUnitCodes = visUnitsString, _
                              Optional ByVal defaultValue As Variant = 0, _
                              Optional ByVal valueToIgnore As Variant) As Collection
'���������� ��������� ���������� �������� � ������� ����� ��������� shpColl
Dim newColl As Collection
Dim tmp As String
Dim shp As Visio.Shape
    
    Set newColl = New Collection
    
    For Each shp In shpColl
        tmp = cellVal(shp, cellName, returnType, defaultValue)
        If valueToIgnore = Empty Then   '���������!
            AddUniqueCollectionItem newColl, tmp, tmp
        Else
            If tmp <> valueToIgnore Then
'                If tmp = "" Then tmp = " "
                AddUniqueCollectionItem newColl, tmp, tmp
            End If
        End If
    Next shp
    
Set GetUniqueVals = newColl
End Function

Public Function StrColToStr(ByRef col As Collection, ByVal delimiter As String) As String
'���������� ������ ��������� �� ��������� �������� ���������
'! �� ��������� �������� ������������
Dim item As Variant
Dim str As String

    On Error Resume Next
    
    str = ""
    For Each item In col
        str = str & item & delimiter
    Next item
    
    str = Left(str, Len(str) - Len(delimiter))
    
StrColToStr = str
End Function

Public Function GetGFSShapeSetTime(ByRef shp As Visio.Shape) As Date
Dim curval As Date

    curval = cellVal(shp, "Prop.ArrivalTime", visDate) + _
                 cellVal(shp, "Prop.LineTime", visDate) + _
                 cellVal(shp, "Prop.SetTime", visDate) + _
                 cellVal(shp, "Prop.SquareTime", visDate) + _
                 cellVal(shp, "Prop.StabCreationTime", visDate) + _
                 cellVal(shp, "Prop.UTPCreationTime", visDate) + _
                 cellVal(shp, "Prop.FormingTime", visDate) + _
                 cellVal(shp, "Prop.FindTime", visDate) + _
                 cellVal(shp, "Prop.RushTime", visDate)
'    Debug.Print curVal

GetGFSShapeSetTime = curval
End Function



Public Sub TTT()
'Debug.Print GetGFSShapeSetTime(Application.ActiveWindow.Selection(1)) < CDate("01.04.2015 11:25:00")
Debug.Print GetGFSShapeSetTime(Application.ActiveWindow.Selection(1)) < 401769
End Sub
'Public Sub TTT()
'Dim c As Collection
'Dim shp As Visio.Shape
'
''Set c = A.Refresh(1).GetGFSShapes("User.IndexPers:" & indexPers.ipStvolRuch & ";User.IndexPers:" & indexPers.ipStvolLafVoda)
'Set c = A.Refresh(1).GetGFSShapes("User.IndexPers:" & indexPers.ipAC)
'Set c = Sort(c, "Prop.ArrivalTime", False)
'
'    For Each shp In c
'        Debug.Print CDate(cellVal(shp, "Prop.ArrivalTime", visDate))
'    Next shp
'
'End Sub


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
