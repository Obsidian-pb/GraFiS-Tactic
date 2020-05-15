Attribute VB_Name = "m_WorkWithlines"
Option Explicit
'----------------------------------Модуль для работы с линиями графиков-----------------------------------



'---------------------------------График расхода----------------------------------------------------------
Public Sub PS_WaterGraph_AddKnot(shp As Visio.Shape)
'Процедура добавляет новый узел к графику изменения расхода воды
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX

    'Блок вставки нового контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count + 1
    PreviosCtrlRowNumber = shp.Section(visSectionControls).Count

    shp.AddNamedRow visSectionControls, "Row_" & CurCtrlRowNumber, visCtlX
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlX).FormulaForceU = "Width*0.95"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlY).FormulaForceU = "Height*0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber & ".Y"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlGlue).FormulaForceU = "TRUE"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlType).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaForceU = """"""


    'Блок создания новых линий и привязки их к новому контролу
    '---Определяем текущие номера линий
    CurLineRowNumber = shp.Section(visSectionFirstComponent).Count
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 1

    '---Привязываем последнюю точку к новому контролу
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 0).FormulaU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 1).FormulaU = "Controls.Row_" & PreviosCtrlRowNumber & ".Y"

    '---Вставляем 1 новую линию
    shp.AddRow visSectionFirstComponent, CurLineRowNumber, visTagLineTo
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 0).FormulaU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"

    '---Вставляем 2 новую линию
    shp.AddRow visSectionFirstComponent, CurLineRowNumber + 1, visTagLineTo
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber + 1, 0).FormulaU = "Width * 1"
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber + 1, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"

    'Блок добавления в массив данных о расходах привязки к новой точке
    '---Для расхода
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_WaterExpenseString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))
    '---Добавляем всплывающую подсказку для контрола
    FormulaString = "Расход: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.A1) & " & Chr(34) & "л/с; Время: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.B1) & " & Chr(34) & "мин."
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaU = """" & FormulaString & """"
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_WaterGraph_AddKnot"
End Sub

Public Sub PS_WaterGraph_DeleteKnot(shp As Visio.Shape)
'Процедура добавляет новый узел к графику изменения расхода воды
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX

    'Блок удаления последнего контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count - 1
    PreviosCtrlRowNumber = shp.Section(visSectionControls).Count - 2
    
    shp.DeleteRow visSectionControls, CurCtrlRowNumber
    
    'Блок удаления последних 2 линий и привязки линии минус 2 к концу
    '---Определяем текущие номера линий
    CurLineRowNumber = shp.Section(visSectionFirstComponent).Count - 1
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 2
    
    shp.DeleteRow visSectionFirstComponent, PreviosLineRowNumber
    shp.DeleteRow visSectionFirstComponent, PreviosLineRowNumber
    
    '---Привязываем последнюю линию к краю фигуры
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber - 1, 0).FormulaU = "Width*1"
    
    'Обновляем массивы данных о точках
    '---Для расхода
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_WaterExpenseString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_WaterGraph_DeleteKnot"
End Sub

'---------------------------------График Площади горения----------------------------------------------------------
Public Sub PS_FireGraph_AddKnot(shp As Visio.Shape)
'Процедура добавляет новый узел к графику изменения площади горения
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX

    'Блок вставки нового контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count + 1
    PreviosCtrlRowNumber = shp.Section(visSectionControls).Count

    shp.AddNamedRow visSectionControls, "Row_" & CurCtrlRowNumber, visCtlX
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlX).FormulaForceU = "Width*(User.EndTime/User.TimeMax)*0.9"  '"Width*0.95"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlY).FormulaForceU = "Height*0.1"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber & ".Y"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlGlue).FormulaForceU = "TRUE"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlType).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaForceU = """"""

    'Блок создания новых линий и привязки их к новому контролу
    '---Определяем текущие номера линий
    CurLineRowNumber = shp.Section(visSectionFirstComponent).Count
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 1

    '---Привязываем последнюю точку к новому контролу
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 0).FormulaU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"

    '---Вставляем 1 новую линию
    shp.AddRow visSectionFirstComponent, CurLineRowNumber, visTagLineTo
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 0).FormulaU = "IF(Actions.NoEndTime.Checked,Width*1,Width*(User.EndTime/User.TimeMax))" '"Width*(User.EndTime/User.TimeMax)"
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 1).FormulaU = "IF(Actions.NoEndTime.Checked,Geometry1.Y" & CurLineRowNumber - 1 & ",Height*0)" ' "Height*0"

    'Блок добавления в массив данных о расходах привязки к новой точке
    '---Для площади
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_FireString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))
    '---Добавляем всплывающую подсказку для контрола
    FormulaString = "Площадь пожара: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.A1) & " & Chr(34) & "м.кв.; Время: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.B1) & " & Chr(34) & "мин."
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaU = """" & FormulaString & """"
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_FireGraph_AddKnot"
End Sub

Public Sub PS_FireGraph_DeleteKnot(shp As Visio.Shape)
'Процедура удаляет последний узел графика изменения ПЛОЩАДИ ГОРЕНИЯ
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX:

    'Блок удаления последнего контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count - 1
    
    shp.DeleteRow visSectionControls, CurCtrlRowNumber
    
    'Блок удаления последних 2 линий и привязки линии минус 2 к концу
    '---Определяем текущие номера линий
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 2
    
    '---Удаляем предпоследнюю строку
    shp.DeleteRow visSectionFirstComponent, PreviosLineRowNumber
    '---Изменяем формулу последней строки
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 1).FormulaU = "IF(Actions.NoEndTime.Checked,Geometry1.Y" & PreviosLineRowNumber - 1 & ",Height*0)"
    
    'Обновляем массивы данных о точках
    '---Для площади
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_FireString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_FireGraph_DeleteKnot"
End Sub

'---------------------------------График Площади пожара----------------------------------------------------------
Public Sub PS_FireTGraph_AddKnot(shp As Visio.Shape)
'Процедура добавляет новый узел к графику изменения площади ПОЖАРА
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX

    'Блок вставки нового контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count + 1
    PreviosCtrlRowNumber = shp.Section(visSectionControls).Count

    shp.AddNamedRow visSectionControls, "Row_" & CurCtrlRowNumber, visCtlX
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlX).FormulaForceU = "Width*0.9"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlY).FormulaForceU = "Controls.Row_" & PreviosCtrlRowNumber & ".Y"  'shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber - 1, visCtlY).Result(visMeters)  '"Height*0.1"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYDyn).FormulaForceU = "Controls.Row_" & CurCtrlRowNumber & ".Y"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlXCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlYCon).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlGlue).FormulaForceU = "TRUE"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlType).FormulaForceU = "0"
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaForceU = """"""

    'Блок создания новых линий и привязки их к новому контролу
    '---Определяем текущие номера линий
    CurLineRowNumber = shp.Section(visSectionFirstComponent).Count
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 1

    '---Привязываем последнюю точку к новому контролу
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 0).FormulaU = "Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"

    '---Вставляем 1 новую линию
    shp.AddRow visSectionFirstComponent, CurLineRowNumber, visTagLineTo
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 0).FormulaU = "Width"  '"Controls.Row_" & CurCtrlRowNumber
    shp.CellsSRC(visSectionFirstComponent, CurLineRowNumber, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"

    'Блок добавления в массив данных о расходах привязки к новой точке
    '---Для площади
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_FireString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))
    '---Добавляем всплывающую подсказку для контрола
    FormulaString = "Площадь пожара: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.A1) & " & Chr(34) & "м.кв.; Время: " & Chr(34) & " & Index(" & PreviosCtrlRowNumber & ", Scratch.B1) & " & Chr(34) & "мин."
    shp.CellsSRC(visSectionControls, PreviosCtrlRowNumber, visCtlTip).FormulaU = """" & FormulaString & """"
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_FireTGraph_AddKnot"
End Sub

Public Sub PS_FireTGraph_DeleteKnot(shp As Visio.Shape)
'Процедура удаляет последний узел графика изменения площади ПОЖАРА
Dim CurCtrlRowNumber As Integer
Dim PreviosCtrlRowNumber As Integer
Dim CurLineRowNumber As Integer
Dim PreviosLineRowNumber As Integer
Dim FormulaString As String

    On Error GoTo EX

    'Блок удаления последнего контрола
    CurCtrlRowNumber = shp.Section(visSectionControls).Count - 1
    
    shp.DeleteRow visSectionControls, CurCtrlRowNumber
    
    'Блок удаления последних 2 линий и привязки линии минус 2 к концу
    '---Определяем текущие номера линий
    PreviosLineRowNumber = shp.Section(visSectionFirstComponent).Count - 2
    
    shp.DeleteRow visSectionFirstComponent, PreviosLineRowNumber
    
    '---Привязываем последнюю линию к контролу перед удаленным
    shp.CellsSRC(visSectionFirstComponent, PreviosLineRowNumber, 1).FormulaU = "Controls.Row_" & CurCtrlRowNumber & ".Y"
    
    'Обновляем массивы данных о точках
    '---Для площади
    shp.CellsSRC(visSectionScratch, 0, visScratchA).FormulaU = pf_FireString(shp.RowCount(visSectionControls))
    '---Для времени
    shp.CellsSRC(visSectionScratch, 0, visScratchB).FormulaU = pf_TimeString(shp.RowCount(visSectionControls))

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_FireTGraph_DeleteKnot"
End Sub



'------------------------------------------ОБЩИЕ---------------------------------------------------------
Public Sub SeekGraphic(ShpObj As Visio.Shape)
'Процедура получения сведений о графике ТП и привязки к нему (для разных типов графиков по-разному)
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim Col As Collection
Dim ShapeType As Integer

On Error GoTo EX

'---Определяем координаты активной фигуры
x = ShpObj.Cells("PinX").Result(visInches)
y = ShpObj.Cells("Piny").Result(visInches)
'---Определяем тип активной фигуры
ShapeType = ShpObj.Cells("User.IndexPers")

'Перебираем все фигуры на странице
For Each OtherShape In Application.ActivePage.Shapes
    If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
        If OtherShape.Cells("User.IndexPers") = 122 And OtherShape.HitTest(x, y, 0.01) > 1 Then  'Если фигура - Новый график ТП
            '---Связываем свойства графика со свойствами поля графика
                ShpObj.Cells("Width").FormulaU = "Sheet." & OtherShape.ID & "!Width"
                ShpObj.Cells("Height").FormulaU = "Sheet." & OtherShape.ID & "!Height"
                ShpObj.Cells("PinX").FormulaU = "Sheet." & OtherShape.ID & "!PinX"
                ShpObj.Cells("PinY").FormulaU = "Sheet." & OtherShape.ID & "!PinY"
                ShpObj.Cells("User.FireMax").FormulaU = "Sheet." & OtherShape.ID & "!Prop.FireMax"
                ShpObj.Cells("User.TimeMax").FormulaU = "Sheet." & OtherShape.ID & "!Prop.TimeMax"
                ShpObj.Cells("LockWidth").Formula = 1
                ShpObj.Cells("LockHeight").Formula = 1
                If ShpObj.CellExists("User.EndTime", 0) = True Then
                    ShpObj.Cells("User.EndTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.TimeEnd"
                End If
                If ShpObj.CellExists("User.Intense", 0) = True Then
                    ShpObj.Cells("User.Intense").FormulaU = "Sheet." & OtherShape.ID & "!User.WaterIntense"
                End If

            '---Для предотвращения дальнейшего исполнения выходим из процедуры
            Set OtherShape = Nothing
            Exit Sub
        End If
    End If
Next OtherShape

'В случае, если ни с одной фигурой поля графика совпадений не найдено, освобождаем привязки
        ShpObj.Cells("Width").FormulaU = ShpObj.Cells("Width")
        ShpObj.Cells("Height").FormulaU = ShpObj.Cells("Height")
        ShpObj.Cells("PinX").FormulaU = ShpObj.Cells("PinX")
        ShpObj.Cells("PinY").FormulaU = ShpObj.Cells("PinY")
        ShpObj.Cells("User.FireMax").FormulaU = ShpObj.Cells("User.FireMax")
        ShpObj.Cells("User.TimeMax").FormulaU = ShpObj.Cells("User.TimeMax")
        ShpObj.Cells("LockWidth").Formula = 0
        ShpObj.Cells("LockHeight").Formula = 0
        If ShpObj.CellExists("User.EndTime", 0) = True Then
            ShpObj.Cells("User.EndTime").FormulaU = ShpObj.Cells("User.EndTime")
        End If
                
ShpObj.BringToFront


Set OtherShape = Nothing
EX:
Set OtherShape = Nothing
End Sub

Private Function pf_WaterExpenseString(ByVal CnotsCount As Integer) As String
'Функция возвращает строку для составления массива данных РАСХОДА ВОДЫ
Dim i As Integer
Dim tmpString As String

tmpString = ""

For i = 1 To CnotsCount
    tmpString = tmpString & "ROUND(User.MaxExpense*Controls.Row_" & i & ".Y/Height,2)&" & Chr(34) & ";" & Chr(34) & "&"
Next i

pf_WaterExpenseString = Left(tmpString, Len(tmpString) - 1)
End Function

Private Function pf_TimeString(ByVal CnotsCount As Integer) As String
'Функция возвращает строку для составления массива данных ВРЕМЕНИ
Dim i As Integer
Dim tmpString As String

tmpString = ""

For i = 1 To CnotsCount
    tmpString = tmpString & "ROUND(User.TimeMax*Controls.Row_" & i & "/Width,2)&" & Chr(34) & ";" & Chr(34) & "&"
Next i

pf_TimeString = Left(tmpString, Len(tmpString) - 1)
End Function

Private Function pf_FireString(ByVal CnotsCount As Integer) As String
'Функция возвращает строку для составления массива данных ПЛОЩАДИ ПОЖАРА
Dim i As Integer
Dim tmpString As String

tmpString = ""

For i = 1 To CnotsCount
    tmpString = tmpString & "ROUND(User.FireMax*Controls.Row_" & i & ".Y/Height,2)&" & Chr(34) & ";" & Chr(34) & "&"
Next i

pf_FireString = Left(tmpString, Len(tmpString) - 1)
End Function
