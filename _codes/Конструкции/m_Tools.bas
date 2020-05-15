Attribute VB_Name = "m_Tools"



Public Function CD_MasterExists(masterName As String) As Boolean
'Функция проверки наличия мастера в активном документе
Dim i As Integer

For i = 1 To Application.ActiveDocument.Masters.count
    If Application.ActiveDocument.Masters(i).Name = masterName Then
        CD_MasterExists = True
        Exit Function
    End If
Next i

CD_MasterExists = False

End Function

Public Sub MasterImportSub(masterName As String)
'Процедура импорта мастера в соответствии с именем
Dim mstr As Visio.Master

    If Not CD_MasterExists(masterName) Then
        Set mstr = ThisDocument.Masters(masterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

End Sub



'-----------------Проки для маски-------------------------------------------------
Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function

Public Function PFB_isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - дверной проем, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isDoor = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ДВЕРЬ или ПРОЕМ
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Or aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25) Then
        PFB_isDoor = True
        Exit Function
    End If
PFB_isDoor = False
End Function

Public Function PFB_isWindow(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - окно, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWindow = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ОКНО
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        aO_Shape.Cells("User.ShapeType").Result(visNumber) = 45 Then
        PFB_isWindow = True
        Exit Function
    End If
PFB_isWindow = False
End Function

'--------------------------------Работа со слоями-------------------------------------
Public Function GetLayerNumber(ByRef layerName As String) As Integer
Dim layer As Visio.layer

    For Each layer In Application.ActivePage.Layers
        If layer.Name = layerName Then
            GetLayerNumber = layer.Index - 1
            Exit Function
        End If
    Next layer
    
    Set layer = Application.ActivePage.Layers.Add(layerName)
    GetLayerNumber = layer.Index - 1
End Function

'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub

'---------------------------------------Служебные функции и проки--------------------------------------------------
Public Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'Функция возвращает угол относительно родительского элемента
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

'Set Shape = Nothing
End Function

Public Sub ClearLayer(ByVal layerName As String)
'Удаляем фигуры указанного слоя
    On Error Resume Next
    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActivePage.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, layerName)
    vsoSelection.Delete
End Sub

Public Function ShapeIsLine(ByRef shp As Visio.Shape) As Boolean
'Функция возвращает истина, если переданная фигура - простая прямая линия, Ложь - если иначе
Dim isLine As Boolean
Dim isStrait As Boolean
    
    ShapeIsLine = False
    
    On Error GoTo EX
    
    If shp.RowCount(visSectionFirstComponent) <> 3 Then Exit Function       'Строк в секции геометрии больше или меньше двух
    If shp.RowType(visSectionFirstComponent, 2) <> 139 Then Exit Function   '139 - LineTo
    
ShapeIsLine = True
Exit Function

EX:
    ShapeIsLine = False
End Function

'--------------------------------------Работа с тулбарами-------------------------------------------------------------
Public Function GetCommandBarTool(ByRef cbr As Office.CommandBar, ByVal toolID As Integer) As Office.CommandBarControl
'Функция возвращает кнопку с указанным ID указанной панели инструментов
Dim btnTool As Office.CommandBarControl
    For Each btnTool In cbr.Controls
        If btnTool.ID = toolID Then
            Set GetCommandBarTool = btnTool
            Exit Function
        End If
    Next btnTool
Set GetCommandBarTool = Nothing
End Function
