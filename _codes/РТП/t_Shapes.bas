Attribute VB_Name = "t_Shapes"
Option Explicit

Public Function cellVal(ByRef shps As Variant, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, Optional defaultValue As Variant = 0) As Variant
'Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
Dim shp As Visio.Shape
Dim tmpVal As Variant
    
    On Error GoTo ex
    
'    Debug.Print TypeName(shps)
    If TypeName(shps) = "Shape" Then        'Если фигура
        Set shp = shps
        If shp.CellExists(cellName, 0) Then
            Select Case dataType
                Case Is = visNumber
                    cellVal = shp.Cells(cellName).Result(dataType)
                Case Is = visUnitsString
                    cellVal = shp.Cells(cellName).ResultStr(dataType)
                Case Is = visDate
                    cellVal = shp.Cells(cellName).Result(dataType)
                Case Else
                    cellVal = shp.Cells(cellName).Result(dataType)
            End Select
        Else
            cellVal = defaultValue
        End If
        Exit Function
    ElseIf TypeName(shps) = "Shapes" Or TypeName(shps) = "Collection" Then     'Если коллекция
        For Each shp In shps
            tmpVal = cellVal(shp, cellName, dataType, defaultValue)
            If tmpVal <> defaultValue Then
                cellVal = tmpVal
                Exit Function
            End If
        Next shp
    End If
    
cellVal = defaultValue
Exit Function
ex:
    cellVal = defaultValue
End Function

Public Sub SetCellFrml(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewFrml As Variant)
'Set cell with cellName formula. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
'    On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        Select Case TypeName(NewFrml)
            Case Is = "String"
                shp.Cells(cellName).FormulaForceU = """" & NewFrml & """"
            Case Is = "Integer"
                shp.Cells(cellName).FormulaForceU = NewFrml
            Case Is = "Long"
                shp.Cells(cellName).FormulaForceU = NewFrml
            Case Is = "Single"
                shp.Cells(cellName).FormulaForceU = NewFrml
            Case Is = "Double"
                shp.Cells(cellName).FormulaForceU = NewFrml
            Case Is = "Date"
                shp.Cells(cellName).FormulaForceU = NewFrml
        End Select
    End If
    
Exit Sub
ex:

End Sub
