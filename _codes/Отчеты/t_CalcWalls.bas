Attribute VB_Name = "t_CalcWalls"
Public Sub PrintWallsLen()
' Расчет суммарной длины стен

Dim TotalLen As Single
Dim wallLen As Single
Dim vsoSelection1 As Visio.Selection
Dim shp As Visio.Shape
Dim item As c_Counter
Dim wallWidth As String
Dim wallsLensCol As Collection

    Set wallsLensCol = New Collection


    TotalLen = 0

    Set vsoSelection1 = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Стена")
    Application.ActiveWindow.Selection = vsoSelection1
    
    For Each shp In vsoSelection1
        wallLen = cellVal(shp, "Width", visMeters)
        
        wallWidth = CStr(cellVal(shp, "Prop.T.Value", visMillimeters))
        If IsKeyInCollection(wallsLensCol, wallWidth) Then
            Set item = GetFromCollection(wallsLensCol, wallWidth)
            item.Rize wallLen
        Else
            Set item = New c_Counter
            item.Activate wallWidth
            item.Rize wallLen
            
            AddUniqueCollectionItem wallsLensCol, item, wallWidth
        End If
        
        
        
        TotalLen = TotalLen + wallLen
    Next shp
    
    Dim resString As String
    resString = "Общая длина стен: " & CStr(Round(TotalLen, 1)) & " м;" & vbNewLine & vbNewLine
    resString = resString & "Из них:" & vbNewLine
    For Each item In wallsLensCol
        resString = resString & item.ID & " мм : " & CStr(Round(item.val, 1)) & " м;" & vbNewLine
    Next item

    MsgBox resString

End Sub

