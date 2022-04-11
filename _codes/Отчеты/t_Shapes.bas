Attribute VB_Name = "t_Shapes"
Public Function IsShapeOnSheet(ByRef shp As Visio.Shape) As Boolean
'Returns True if shape is in page rect
Dim X As Double
Dim Y As Double

    X = shp.Cells("PinX").Result(visInches)
    Y = shp.Cells("PinY").Result(visInches)

    If X < 0 Or X > Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches) Or _
        Y < 0 Or Y > Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches) Then
        IsShapeOnSheet = False
        Exit Function
    End If
    
IsShapeOnSheet = True
End Function

'Public Function cellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, _
'                        Optional ByVal defaultValue As Double = 0) As Variant
''Returns cell with cellName value. If such cell does not exists, return 0
'
'    On Error GoTo EX
'
'    If shp.CellExists(cellName, 0) Then
'        Select Case dataType
'            Case Is = visNumber
'                cellVal = shp.Cells(cellName).Result(dataType)
'            Case Is = visUnitsString
'                cellVal = shp.Cells(cellName).ResultStr(dataType)
'            Case Is = visDate
'                cellVal = shp.Cells(cellName).Result(dataType)
'            Case Else
'                cellVal = shp.Cells(cellName).Result(dataType)
'        End Select
'    Else
'        cellVal = defaultValue
'    End If
'
'Exit Function
'EX:
'    cellVal = defaultValue
'End Function
'!!!Рабочая
'Public Function cellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, Optional defaultValue As Variant = 0) As Variant
''Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
'
'    On Error GoTo EX
'
'    If shp.CellExists(cellName, 0) Then
'        Select Case dataType
'            Case Is = visNumber
'                cellVal = shp.Cells(cellName).Result(dataType)
'            Case Is = visUnitsString
'                cellVal = shp.Cells(cellName).ResultStr(dataType)
'            Case Is = visDate
'                cellVal = shp.Cells(cellName).Result(dataType)
'            Case Else
'                cellVal = shp.Cells(cellName).Result(dataType)
'        End Select
'    Else
'        cellVal = defaultValue
'    End If
'
'
'Exit Function
'EX:
'    cellVal = defaultValue
'End Function
Public Function cellVal(ByRef shps As Variant, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, Optional defaultValue As Variant = 0) As Variant
'Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
Dim shp As Visio.Shape
Dim tmpVal As Variant
    
    On Error GoTo EX
    
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
'                    Debug.Assert cellName <> "Prop.FireTime"
'                    If cellName = "Prop.FireTime" Then Stop
                    If cellVal = 0 Then
                        cellVal = CDate(shp.Cells(cellName).ResultStr(visUnitsString))
                    End If
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
EX:
    cellVal = defaultValue
End Function





Public Sub SetCellVal(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewVal As Variant)
'Set cell with cellName value. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForce = """" & NewVal & """"
    End If
    
Exit Sub
EX:
    Debug.Print "Error in t_Shapes modul in 'Otcheti'! " & shp.Name & ", " & cellName & ", " & NewVal
End Sub

Public Sub SetCellFrml(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewFrml As Variant)
'Set cell with cellName formula. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForceU = NewFrml
    End If
    
Exit Sub
EX:

End Sub

Public Function ShapeHaveCell(ByRef shp As Visio.Shape, ByVal cellName As String, _
                              Optional ByVal val As Variant = "", Optional ByVal delimiter As Variant = ";") As Boolean
'Функция возвращает Истина, если такая ячейка есть
'Если указано значение ячейки, также проверяется и она. Если указано несколько значений, то проверяются все они
Dim vals() As String
Dim curval As String
Dim i As Integer

On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        If val <> "" Then
            ' Проверяем одно ли значение передано в атрибуте val
            If InStr(1, val, delimiter) > 0 Then
                '---Если значений несколько
                vals = Split(val, delimiter)
                For i = 0 To UBound(vals)
                    curval = vals(i)
                    If shp.Cells(cellName).ResultStr(visUnitsString) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    ElseIf shp.Cells(cellName).Result(visNumber) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    Else
                        ShapeHaveCell = False
                    End If
                Next i
            Else
                '---Если значение одно
                If shp.Cells(cellName).ResultStr(visUnitsString) = val Then
                    ShapeHaveCell = True
                ElseIf shp.Cells(cellName).Result(visNumber) = val Then
                    ShapeHaveCell = True
                Else
                    ShapeHaveCell = False
                End If
            End If
        Else
            ShapeHaveCell = True
        End If
    Else
        ShapeHaveCell = False
    End If
    
Exit Function
EX:
    ShapeHaveCell = False
End Function





