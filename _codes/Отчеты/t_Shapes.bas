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
    
'    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForce = """" & NewVal & """"
    End If
    
Exit Sub
EX:

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
                              Optional ByVal val As Variant = "") As Boolean
On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        If val <> "" Then
            If shp.Cells(cellName).ResultStr(visUnitsString) = val Then
                ShapeHaveCell = True
            ElseIf shp.Cells(cellName).Result(visNumber) = val Then
                ShapeHaveCell = True
            Else
                ShapeHaveCell = False
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

'Public Sub FixLineGroupProportions()
''Main sub for fixation the line width of each shape in groupe
'Dim shp As Visio.Shape
'
'    Set shp = Application.ActiveWindow.Selection(1)
'
'    FixLineProportions shp
'End Sub
'
'Public Sub FixLineProportions(ByRef shp As Visio.Shape)
''Fixes the line width of each shape in groupe
'Dim lineProp As Double
'Dim frml As String
'Dim shp2 As Visio.Shape
'
'    lineProp = cellVal(shp, "LineWeight") / cellVal(shp, "Height")
'
'    frml = "Height*" & lineProp & "*(ThePage!PageScale/ThePage!DrawingScale)"
'    shp.Cells("LineWeight").Formula = frml
'
'    If shp.Shapes.count > 0 Then
'        For Each shp2 In shp.Shapes
'            FixLineProportions shp2
'        Next shp2
'    End If
'
'End Sub

'Public Sub FixTextGroupProportions()
''Main sub for fixation the text height of each shape in groupe
'Dim shp As Visio.Shape
'
'    Set shp = Application.ActiveWindow.Selection(1)
'
'    FixTextProportions shp
'End Sub
'
'Public Sub FixTextProportions(ByRef shp As Visio.Shape)
''Fixes the text height of each shape in groupe
'Dim textProp As Double
'Dim frml As String
'Dim shp2 As Visio.Shape
'
'    textProp = cellVal(shp, "Char.Size") / cellVal(shp, "Height")
'
'    frml = "Height*" & textProp & "*(ThePage!PageScale/ThePage!DrawingScale)"
'    shp.Cells("Char.Size").Formula = frml
'    shp.Cells("LeftMargin").Formula = 0
'    shp.Cells("RightMargin").Formula = 0
'    shp.Cells("TopMargin").Formula = 0
'    shp.Cells("BottomMargin").Formula = 0
'
'    If shp.Shapes.count > 0 Then
'        For Each shp2 In shp.Shapes
'            FixTextProportions shp2
'        Next shp2
'    End If
'
'End Sub

'Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPerses As Variant, Optional needGFSChecj As Boolean = False) As Boolean
''Функция возвращает True, если фигура является фигурой ГраФиС и среди переданных типов фигур ГраФиС (gfsIndexPreses) присутствует IndexPers данной фигуры
''По умолчанию предполагается что переданная фигура уже проверена на то, относится ли она к фигурам ГраФиС. В случае, если у фигуры нет ячейки User.IndexPers _
''обработчик ошибки указывает функции вернуть False
''Пример использования: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
''                 или: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
'Dim i As Integer
'Dim indexPers As Integer
'
'    On Error GoTo EX
'
'    'Если необходима предварительная проверка на отношение фигуры к ГраФиС:
'    If needGFSChecj Then
'        If Not IsGFSShape(shp) Then
'            IsGFSShapeWithIP = False
'            Exit Function
'        End If
'    End If
'
'    'Проверяем, является ли фигура фигурой указанного типа
'    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
'    Select Case TypeName(gfsIndexPerses)
'        Case Is = "Long"    'Если передано единственное значение Long
'            If gfsIndexPerses = indexPers Then
'                IsGFSShapeWithIP = True
'                Exit Function
'            End If
'        Case Is = "Integer"    'Если передано единственное значение Integer
'            If gfsIndexPerses = indexPers Then
'                IsGFSShapeWithIP = True
'                Exit Function
'            End If
'        Case Is = "Variant()"   'Если передан массив
'            For i = 0 To UBound(gfsIndexPerses)
'                If gfsIndexPerses(i) = indexPers Then
'                    IsGFSShapeWithIP = True
'                    Exit Function
'                End If
'            Next i
'        Case Else
'            IsGFSShapeWithIP = False
'    End Select
'
'IsGFSShapeWithIP = False
'Exit Function
'EX:
'    IsGFSShapeWithIP = False
'    SaveLog Err, "m_Tools.IsGFSShapeWithIP"
'End Function


