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

Public Function GetGFSShapes(ByRef cellArr As Variant, Optional ByVal delimiter As String = ":") As Collection
'Применение: GetGFSShapes(Array("User.IndexPers:500","User.IndexPers:1"))
Dim shp As Visio.Shape
Dim shpCellItems() As String
Dim i As Integer
Dim tmpColl As Collection
    
    Set tmpColl = New Collection
    
    For Each shp In Application.ActivePage.Shapes
        For i = 0 To UBound(cellArr)
            shpCellItems = Split(cellArr(i), delimiter)
            If ShapeHaveCell(shp, shpCellItems(0), shpCellItems(1)) Then
'                tmpColl.Add shp
                AddUniqueCollectionItem tmpColl, shp
            End If
        Next i
    Next shp
    
    If tmpColl.Count > 0 Then
        Set GetGFSShapes = tmpColl
    Else
        Set GetGFSShapes = Nothing
    End If
End Function
