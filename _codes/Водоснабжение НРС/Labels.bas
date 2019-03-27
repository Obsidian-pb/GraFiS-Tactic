Attribute VB_Name = "Labels"
Sub InsertLabelName(ShpObj As Visio.Shape)
'Процедура добавления подписи названия открытого водоисточника связанной с фигурой водоисточника
'---Объявляем переменные
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.cell, Cell2 As Visio.cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX

'---Вбрасываем фигуру подписи площади
    '---Определяем точки Х и У для вброса
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Водоснабжение НРС.vss").Masters.ItemU("Подпись названия водоисточника"), pntX, pntY)

'---Вбрасываем коннектор и соединяем фигуру водоисточника и подпись
    '---Вбрасываем коннектор и соединяем фигкуры водоисточника и подписи
    Set mstrConnection = Application.Documents("Водоснабжение НРС.vss").Masters("Коннектор")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set vsoCell1 = shpConnection.CellsU("BeginX")
    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    Set vsoCell1 = shpConnection.CellsU("EndX")
    Set vsoCell2 = shpLabel.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2

'---Определяем свойства фигуры коннектора
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(80,123,175))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---Добавляем связь подписи с фигурой водоисточника
    CellFormula = "Sheet." & ShpObj.ID & "!Prop.Name"
    shpLabel.CellsSRC(visSectionTextField, 0, visFieldCell).FormulaU = CellFormula

   
'---Ставим фокус на подписи
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelName"
End Sub

Sub InsertLabelValue(ShpObj As Visio.Shape)
'Процедура добавления подписи объема открытого водоисточника связанной с фигурой водоисточника
'---Объявляем переменные
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.cell, Cell2 As Visio.cell
Dim CellFormula As String
'Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX
'---Вбрасываем фигуру объема вобоисточника
    '---Определяем точки Х и У для вброса
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.Item("Водоснабжение НРС.vss").Masters.ItemU("Объем открытого водоисточника"), pntX, pntY)

'---Вбрасываем коннектор и соединяем фигуру водоисточника и подпись
    '---Вбрасываем коннектор и соединяем фигкуры водоисточника и подписи
    Set mstrConnection = Application.Documents("Водоснабжение НРС.vss").Masters("Коннектор")
    
    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
    Set vsoCell1 = shpConnection.CellsU("BeginX")
    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2
    Set vsoCell1 = shpConnection.CellsU("EndX")
    Set vsoCell2 = shpLabel.CellsSRC(1, 1, 0)
        vsoCell1.GlueTo vsoCell2

'---Определяем свойства фигуры коннектора
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLOLineRouteExt).FormulaU = 1
    shpConnection.CellsSRC(visSectionObject, visRowShapeLayout, visSLORouteStyle).FormulaU = 16
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(80,123,175))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---Добавляем связь подписи с фигурой водоисточника
    CellFormula = "Sheet." & ShpObj.ID & "!Prop.Value"
    shpLabel.CellsSRC(visSectionTextField, 0, visFieldCell).FormulaU = CellFormula

   
'---Ставим фокус на подписи
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelValue"
End Sub

Public Sub ConnectedShapesLostCheck(ShpObj As Visio.Shape)
'Процедура проверяет, не была ли удалена одна из фигур соединенных коннектором, и если была, то удаляет сам коннектор
Dim CellsVal(4) As String
    
On Error GoTo EX
    
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
