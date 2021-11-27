Attribute VB_Name = "t_Shapes"
Public Function IsShapeOnSheet(ByRef shp As Visio.Shape) As Boolean
'Returns True if shape is in page rect
Dim x As Double
Dim y As Double

    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)

    If x < 0 Or x > Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches) Or _
        y < 0 Or y > Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches) Then
        IsShapeOnSheet = False
        Exit Function
    End If
    
IsShapeOnSheet = True
End Function

Public Function CellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, _
                        Optional ByVal defaultValue As Double = 0) As Variant
'Returns cell with cellName value. If such cell does not exists, return 0
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        Select Case dataType
            Case Is = visNumber
                CellVal = shp.Cells(cellName).Result(dataType)
            Case Is = visUnitsString
                CellVal = shp.Cells(cellName).ResultStr(dataType)
            Case Is = visDate
                CellVal = shp.Cells(cellName).Result(dataType)
            Case Else
                CellVal = shp.Cells(cellName).Result(dataType)
        End Select
    Else
        CellVal = defaultValue
    End If
    
Exit Function
ex:
    CellVal = defaultValue
End Function

Public Sub SetCellVal(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewVal As Variant)
'Set cell with cellName value. If such cell does not exists, does nothing
Dim cll As Visio.cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForce = NewVal
    End If
    
Exit Sub
ex:

End Sub

Public Sub SetCellFrml(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewFrml As Variant)
'Set cell with cellName formula. If such cell does not exists, does nothing
Dim cll As Visio.cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForceU = NewFrml
    End If
    
Exit Sub
ex:

End Sub

Public Sub FixLineGroupProportions()
'Main sub for fixation the line width of each shape in groupe
Dim shp As Visio.Shape
    
    Set shp = Application.ActiveWindow.Selection(1)
    
    FixLineProportions shp
End Sub

Public Sub FixLineProportions(ByRef shp As Visio.Shape)
'Fixes the line width of each shape in groupe
Dim lineProp As Double
Dim frml As String
Dim shp2 As Visio.Shape
    
    lineProp = CellVal(shp, "LineWeight") / CellVal(shp, "Height")
    
    frml = "Height*" & lineProp & "*(ThePage!PageScale/ThePage!DrawingScale)"
    shp.Cells("LineWeight").Formula = frml
    
    If shp.Shapes.Count > 0 Then
        For Each shp2 In shp.Shapes
            FixLineProportions shp2
        Next shp2
    End If
    
End Sub

Public Sub FixTextGroupProportions()
'Main sub for fixation the text height of each shape in groupe
Dim shp As Visio.Shape
    
    Set shp = Application.ActiveWindow.Selection(1)
    
    FixTextProportions shp
End Sub

Public Sub FixTextProportions(ByRef shp As Visio.Shape)
'Fixes the text height of each shape in groupe
Dim textProp As Double
Dim frml As String
Dim shp2 As Visio.Shape
    
    textProp = CellVal(shp, "Char.Size") / CellVal(shp, "Height")
    
    frml = "Height*" & textProp & "*(ThePage!PageScale/ThePage!DrawingScale)"
    shp.Cells("Char.Size").Formula = frml
    shp.Cells("LeftMargin").Formula = 0
    shp.Cells("RightMargin").Formula = 0
    shp.Cells("TopMargin").Formula = 0
    shp.Cells("BottomMargin").Formula = 0
    
    If shp.Shapes.Count > 0 Then
        For Each shp2 In shp.Shapes
            FixTextProportions shp2
        Next shp2
    End If
    
End Sub
