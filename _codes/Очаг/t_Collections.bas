Attribute VB_Name = "t_Collections"
Option Explicit

'--------------------------------���������-------------------------------------
Public Sub AddCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'��������� �������� ����� ��������� � ������
Dim GenPointItem As c_Point

    '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
    For Each GenPointItem In newCollection
        oldCollection.Add GenPointItem
    Next GenPointItem
End Sub

Public Sub AddUniqueCollectionItems(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'��������� ����� �������� ����� ��������� � ������
Dim GenPointItem As c_Point

    '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
    For Each GenPointItem In newCollection
       AddUniqueCollectionItem oldCollection, GenPointItem
    Next GenPointItem
End Sub

'Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As c_Point)
Public Sub AddUniqueCollectionItem(ByRef oldCollection As Collection, ByRef item As Variant)
Dim GenPointItem As c_Point
    
    If TypeName(item) = "c_Point" Then
        '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
        For Each GenPointItem In oldCollection
            If GenPointItem.x = item.x And GenPointItem.y = item.y Then Exit Sub
        Next GenPointItem
    
        oldCollection.Add item
    ElseIf TypeName(item) = "Shape" Then
        On Error Resume Next
        oldCollection.Add item, CStr(item.ID)
    End If
End Sub

Public Sub SetCollection(ByRef oldCollection As Collection, ByRef newCollection As Collection)
'��������� ������ ��������� � ������������ �� ���������� ����� ���������
Dim item As Variant

    Set oldCollection = New Collection

    '---���������� ��� �������� � ����� ��������� � ��������� �� � ������
    For Each item In newCollection
        oldCollection.Add item
    Next item

    '---������� ����� ���������
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



Public Function FilterShapes(ByRef shpColl As Variant, ByVal filterStr As String, _
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
    
    On Error GoTo EX
    
    Set tmpColl = New Collection
    
    filters = Split(filterStr, d_elem)
    
'    Debug.Print TypeName(shpColl)
    For Each shp In shpColl
        '�������� �� shp!
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

Public Function TryGetShape(ByRef shp As Visio.Shape, ByVal filterStr As String, _
                         Optional ByVal d_elem As String = ";", Optional ByVal d_val As String = ":") As Boolean
'�������� �������� ������ ������ ������ ������ ��������������� ������� filterStr
Dim shpTmp As Visio.Shape
Dim col As Collection
    
    On Error GoTo EX
    
    Set col = FilterShapes(Application.ActivePage.Shapes, filterStr, d_elem, d_val)
    If col.Count >= 1 Then
        Set shp = col(1)
        TryGetShape = True
    Else
        TryGetShape = False
    End If
    
Exit Function
EX:
    TryGetShape = False
End Function

'Public Sub Test1()
'Dim shp As Visio.Shape
'
'    Debug.Print TryGetShape(shp, "User.IndexPers:1001")
'
'    Debug.Print shp.ID
'End Sub

