Attribute VB_Name = "m_Editor"
Option Explicit




Public Sub SetDoorAsExit()
'Устанавливаем для выбранных дверей свойство "ВЫХОД"
Dim shp As Visio.Shape
Dim rowI As Integer
    
    For Each shp In Application.ActiveWindow.Selection
        If PFB_isDoor(shp) Then
            Debug.Print shp.Name
            With shp
                If Not .CellExists("User.DoorIsExit", 0) Then
                    rowI = .AddNamedRow(visSectionUser, "DoorIsExit", 0)
                Else
                    rowI = .Cells("User.DoorIsExit").Row
                End If
                .Cells("User.DoorIsExit").Formula = 1
                .Cells("User.DoorIsExit.Prompt").Formula = """" & "This door is exit from floor" & """"
            End With
            setShpColor shp, "3"
        End If
    Next shp
End Sub

Private Sub setShpColor(ByRef shp As Visio.Shape, ByVal color As String)
Dim shpChild As Visio.Shape

    shp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = color
    
    If shp.Shapes.count > 0 Then
        For Each shpChild In shp.Shapes
            setShpColor shpChild, color
        Next shpChild
    End If
End Sub

