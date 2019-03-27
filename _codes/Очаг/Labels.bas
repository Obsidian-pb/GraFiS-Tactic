Attribute VB_Name = "Labels"
Sub InsertLabelSquare(ShpObj As Visio.Shape)
'Процедура добавления подписи площади пожара связанной с фигурой площади
'---Объявляем переменные
Dim shpLabel As Visio.Shape
Dim shpConnection As Visio.Shape
Dim mstrConnection As Visio.Master
Dim Cell1 As Visio.cell, Cell2 As Visio.cell
Dim CellFormula As String
Dim pnt1 As Long, pnt2 As Long

    On Error GoTo EX
'---Вбрасываем фигуру подписи площади
    '---Определяем точки Х и У для вброса
        pntX = ShpObj.CellsU("pinX")
        pntY = ShpObj.CellsU("pinY")
    Set shpLabel = Application.ActiveWindow.Page.Drop(Application.Documents.item("Очаг.vss").Masters.ItemU("Подпись площадь"), pntX, pntY)

'---Вбрасываем коннектор и соединяем фигуру площади и подпись
    '---Вбрасываем коннектор и соединяем фигкуры площади и подписи
    Set mstrConnection = Application.Documents("Очаг.vss").Masters("Коннектор")
    
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
    shpConnection.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(255,0,0))"

    CellFormula = "AND(EndX>Sheet." & ShpObj.ID & "!PinX-Sheet." & ShpObj.ID & "!Width*0.5,EndX<Sheet." & _
        ShpObj.ID & "!PinX+Sheet." & ShpObj.ID & "!Width*0.5,EndY<Sheet." & _
        ShpObj.ID & "!PinY+Sheet." & ShpObj.ID & "!Height*0.5,EndY>Sheet." & _
        ShpObj.ID & "!PinY-Sheet." & ShpObj.ID & "!Height*0.5)"
    shpConnection.CellsSRC(visSectionFirstComponent, 0, 1).FormulaU = CellFormula

'---Добавляем связь подписи с фигурой площади
    CellFormula = "Sheet." & ShpObj.ID & "!User.FireSquare"
    shpLabel.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = CellFormula
    CellFormula = "Sheet." & ShpObj.ID & "!User.ExtSquare"
    shpLabel.CellsSRC(visSectionUser, 1, visUserValue).FormulaU = CellFormula
    
'---Скрываем поле стандартного значения
    shpLabel.Cells("Prop.Square.Invisible").FormulaU = 1
   
'---Ставим фокус на подписи
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select shpLabel, visSelect
    
'---Показываем свойства
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
EX:
    SaveLog Err, "InsertLabelSquare"
End Sub


Public Sub SeekFire(ShpObj As Visio.Shape)
'Процедура получения скорости распространения огня и присвоение его фигуре направления распространения огня
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim col As Collection

    On Error GoTo EX
'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'Перебираем все фигуры на странице
    For Each OtherShape In Application.ActivePage.Shapes
        If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.Cells("User.IndexPers") = 64 And OtherShape.HitTest(x, y, 0.01) > 1 Then
                ShpObj.Cells("Prop.FireSpeed").FormulaU = _
                 "Sheet." & OtherShape.ID & "!Prop.FireSpeedLine"
            End If
        End If
    Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    SaveLog Err, "SeekFire", ShpObj.Name
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

'Public Sub HideMaster()
'Dim mstrConnection As Visio.Master
'
'    Set mstrConnection = Application.Documents("Водоснабжение.vss").Masters("Коннектор")
'
'    mstrConnection.Hidden = True
'
'End Sub
