Attribute VB_Name = "m_WorkWithUniversalLabels"
Option Explicit


'------------------------Модуль для хранения процедур работы с универсальными подписями----------


Public Sub SeekAnyGFSFigure(ShpObj As Visio.Shape)
'Публичная прока проверяет все фигуры на листе и если эта фигура, на которую вброшена
'универсальная подпись, связываем их
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim str As String
Dim delFlag As Boolean

    On Error GoTo EX
'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---Перебираем все фигуры на странице
    delFlag = True
    For Each OtherShape In Application.ActivePage.Shapes
        If IsCorrectShapeForLabel(OtherShape) Then 'And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.HitTest(x, y, 10) >= 1 Then
            '---Устанавливаем содеримое ячейки Prop.Property (перечень свойств)
                ShpObj.Cells("Prop.Property.Format").Formula = """" & GetPropsList(OtherShape) & """" '"Пользовательская"
            '---Устанавливаем ссылки на ячейки Prop.Property фигуры
                str = GetPropsLinks(OtherShape)
                ShpObj.Cells("Prop.PropertyValue.Format").Formula = str
                
            '---Вбрасываем соединительную линию
                InsertLink OtherShape, ShpObj
                
            '---Указываем что удалять фигуру после отрисовки не нужно
                delFlag = False
                
            '---Выходим из цикла
                Exit For
            End If
        End If
    Next OtherShape

'---В случае, если фигура не была ни к чему приклеена - удаляем ее
    If delFlag Then
        ShpObj.Delete
        Exit Sub
    End If
    
'---Показываем свойства
    On Error Resume Next
    Application.DoCmd (1312)

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    MsgBox "В ходе выполнения произошла ошибка, если она будет повторяться - свяжитесь с разработчиком!"
    SaveLog Err, "SeekAnyGFSFigure", ShpObj.Name
End Sub

Public Sub ConnectToGFSFigure(ShpObj As Visio.Shape)
'Публичная прока проверяет к какой фигуре была приклеена
'универсальная подпись, связывает их
Dim OtherShape As Visio.Shape
Dim str As String

    On Error GoTo EX
    
'---Определяем сколько соединений у фигуры универсальной подписи, и если ровно 1 соединяем фигуры
    If ShpObj.Connects.Count = 1 Then
        Set OtherShape = ShpObj.Connects.Item(1).ToSheet
        GetCorrectShape OtherShape
    
    '---Устанавливаем содеримое ячейки Prop.Property (перечень свойств)
        ShpObj.Cells("Prop.Property.Format").Formula = """" & GetPropsList(OtherShape) & """" '"Пользовательская"
    '---Устанавливаем ссылки на ячейки Prop.Property фигуры
        str = GetPropsLinks(OtherShape)
        ShpObj.Cells("Prop.PropertyValue.Format").Formula = str
        
    Else
    '---В случае, если кол-во соединений не равно 1 ставим значения подписей по умолчанию
        ShpObj.Cells("Prop.Property.Format").Formula = """" & "Пользовательская;Первая" & """"
        ShpObj.Cells("Prop.PropertyValue.Format").Formula = "Prop.Property.Prompt&" & Chr(34) & ";" & Chr(34) & _
            "&Prop.UserLabel.Prompt"
    End If

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    MsgBox "В ходе выполнения произошла ошибка, если она будет повторяться - свяжитесь с разработчиком!"
    SaveLog Err, "ConnectToGFSFigure", ShpObj.Name
End Sub

Private Function IsCorrectShapeForLabel(ShpTO As Visio.Shape)
'Функция проверяет, можно ли прилепить подпись к данной фигуре
IsCorrectShapeForLabel = True
'---проверка на фигуру подписи
    If ShpTO.CellExists("User.IndexPers", 0) = True Then
        If ShpTO.Cells("User.IndexPers").Result(visNumber) = 152 Then
            IsCorrectShapeForLabel = False
        End If
    End If
'---проверка на фигуру линий (1D)
    If InStr(1, ShpTO.Cells("Width").FormulaU, "SQRT") > 0 Then 'Корнем длина вычисляется только для линий, т.е. если есть SQRT - то эта фигура точно линия 1D!!!
        IsCorrectShapeForLabel = False
    End If
'---Проверка на коннектор
    If Not ShpTO.Cells("BegTrigger").FormulaU = "" Or Not ShpTO.Cells("EndTrigger").FormulaU = "" Then 'У коннекторов эти ячейки всегда заполнены!!!
        IsCorrectShapeForLabel = False
    End If
'---Проверка наличия свойств
    If ShpTO.SectionExists(visSectionProp, 0) = False Then
        IsCorrectShapeForLabel = False
    End If
    
End Function

Sub InsertLink(ShpTO As Visio.Shape, ShpFROM As Visio.Shape)
'Процедура добавления подписи площади пожара связанной с фигурой площади
'---Объявляем переменные
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.Cell, Cell2 As Visio.Cell
Dim CellFormula As String
Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX

'---Вбрасываем коннектор и соединяем фигуру площади и подпись
    '---Вбрасываем коннектор и соединяем фигкуры площади и подписи
    Set mstrConnection = ThisDocument.Masters("Коннектор")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set Cell1 = shpConnection.CellsU("BeginX")
    Set Cell2 = ShpTO.CellsSRC(1, 1, 0)
        Cell1.GlueTo Cell2
    Set Cell1 = shpConnection.CellsU("EndX")
    Set Cell2 = ShpFROM.CellsSRC(1, 1, 0)
        Cell1.GlueTo Cell2
        
        
    '---Задаем привязку к ConnectionPoints фигуры подпси
    CellFormula = "IF(BeginX<Sheet." & ShpFROM.ID & "!PinX,PAR(PNT(Sheet." & ShpFROM.ID & _
        "!Connections.LeftConPoint.X,Sheet." & ShpFROM.ID & "!Connections.LeftConPoint.Y)),PAR(PNT(Sheet." & ShpFROM.ID & _
        "!Connections.RIghtConPoint.X,Sheet." & ShpFROM.ID & "!Connections.RIghtConPoint.Y)))"
        shpConnection.CellsU("EndX").FormulaU = CellFormula
        shpConnection.CellsU("EndY").FormulaU = CellFormula
    
'---Определяем свойства фигуры коннектора
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    
    CellFormula = "AND(EndX>Sheet." & ShpTO.ID & "!PinX-Sheet." & ShpTO.ID & "!Width*0.5,EndX<Sheet." & _
        ShpTO.ID & "!PinX+Sheet." & ShpTO.ID & "!Width*0.5,EndY<Sheet." & _
        ShpTO.ID & "!PinY+Sheet." & ShpTO.ID & "!Height*0.5,EndY>Sheet." & _
        ShpTO.ID & "!PinY-Sheet." & ShpTO.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula
   
   CellFormula = "Sheet." & ShpFROM.ID & "!LineColor"
    shpConnection.Cells("LineColor").FormulaU = CellFormula
   
'---Ставим фокус на подписи
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpFROM, visSelect
    

    
Exit Sub
EX:
'    MsgBox "В ходе выполнения произошла ошибка, если она будет повторяться - свяжитесь с разработчиком!"
    SaveLog Err, "InsertLabelSquare"
'---Ставим фокус на подписи
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpFROM, visSelect
'---Удаляем коннектор
    shpConnection.Delete
End Sub


Private Function GetPropsList(ByRef DirShpObj As Visio.Shape) As String
'Получаем перечень имен свойств указанной фигуры
Dim i As Integer
Dim tempStr As String
    
    tempStr = "Пользовательская"
    If DirShpObj.SectionExists(visSectionProp, 0) = True Then
    '---Если секция есть - заполняем списко
        For i = 0 To DirShpObj.RowCount(visSectionProp) - 1
            If DirShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).Result(visNone) = 0 Then
                tempStr = tempStr & ";" & DirShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone)
            End If
        Next i
    End If

GetPropsList = tempStr
End Function

Private Function GetPropsLinks(ByRef DirShpObj As Visio.Shape) As String
'Получаем перечень ссылок на свойства указанной фигуры
Dim i As Integer
Dim tempStr As String
    
    tempStr = "Prop.Property.Prompt"
    If DirShpObj.SectionExists(visSectionProp, 0) = True Then
    '---Если секция есть - заполняем списко
        For i = 0 To DirShpObj.RowCount(visSectionProp) - 1
            If DirShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).Result(visNone) = 0 Then
                tempStr = tempStr & Chr(38) & Chr(34) & ";" & Chr(34) & Chr(38) & _
                            "Sheet." & DirShpObj.ID & "!" & _
                            DirShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).Name
            End If
        Next i
    End If

GetPropsLinks = tempStr

End Function

Private Sub GetCorrectShape(ByRef shp As Visio.Shape)
Dim corrShp As Visio.Shape
    
    On Error GoTo EX
    Set corrShp = shp.Parent
    '---Если родительская фигура != Страница
    GetCorrectShape corrShp
    Set shp = corrShp
    
Exit Sub
EX:
End Sub


Public Sub ConnectedShapesLostCheck(ShpObj As Visio.Shape)
'Процедура проверяет, не была ли удалена одна из фигур соединенных коннектором, и если была, то удаляет сам коннектор
Dim CellsVal(4) As String
    
'    If ShpObj Is Nothing Then MsgBox "12"
    
On Error GoTo EX
    
'    Debug.Print ShpObj.Name
    
    CellsVal(0) = ShpObj.Cells("BegTrigger").FormulaU
    CellsVal(1) = ShpObj.Cells("BegTrigger").Result(visUnitsString)
    CellsVal(2) = ShpObj.Cells("EndTrigger").FormulaU
    CellsVal(3) = ShpObj.Cells("EndTrigger").Result(visUnitsString)
    
    If CellsVal(0) = CellsVal(1) Or CellsVal(2) = CellsVal(3) Then
        ShpObj.Delete
    End If
    

Exit Sub
EX:
    'Ошибка
End Sub


