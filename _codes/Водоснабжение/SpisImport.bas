Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)

'---Обновляем общие списки
    ShpObj.Cells("Prop.PipeType.Format").FormulaU = ListImport("Вид сети", "Вид водовода")
    
'---Обновляем список моделей и их ТТХ
    '---Запускаем процедуру получения СПИСКА диаметров
    DiametersListImport (ShpObj.ID)
    '---Запускаем процедуру получения СПИСКА напоров
    PressuresListImport (ShpObj.ID)
    '---Запускаем процедуру пересчета водоотдачи
    ProductionImport (ShpObj.ID)

On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШБИКИ
Application.DoCmd (1312)

End Sub

'------------------------Блок зависимых списков------------------------------

Public Sub DiametersListImport(ShpIndex As Long)
'Процедура импорта списков диаметров
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
    Select Case IndexPers
        Case Is = 50 'Пожарный гидрант
            Criteria = "[Вид водовода] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' "
            shp.Cells("Prop.PipeDiameter.Format").FormulaU = ListImportNum("ЗапросВодоотдачи", "Диаметр водовода", Criteria)
    
    End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
'    If shp.Cells("Prop.PipeDiameter").ResultStr(Visio.visNone) = "" Then
'        shp.Cells("Prop.PipeDiameter").FormulaU = "INDEX(0,Prop.PipeDiameter.Format)"
'    End If
    
    Set shp = Nothing

Exit Sub
EX:
    SaveLog Err, "DiametersListImport", CStr(ShpIndex)
End Sub


Public Sub PressuresListImport(ShpIndex As Long)
'Процедура импорта списка возможных напоров для заданных условий
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка возможных струй стволов для текущей фигуры
    Select Case IndexPers
        Case Is = 50
            Criteria = "[Вид водовода] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' And " & _
                "[Диаметр водовода] = " & shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.Pressure.Format").FormulaU = ListImportNum("ЗапросВодоотдачи", "Напор в сети", Criteria)
    End Select

'---В случае, если значение поля для нового списка равно "", включаем пользовательский ввод
    If shp.Cells("Prop.Pressure").ResultStr(Visio.visNone) = "" Then
        'shp.Cells("Prop.Pressure").FormulaU = "INDEX(0,Prop.Pressure.Format)"
        shp.Cells("Prop.ShowDirectProduction").FormulaU = "INDEX(1,Prop.ShowDirectProduction.Format)"
    End If

    Set shp = Nothing

Exit Sub
EX:
    SaveLog Err, "PressuresListImport", ShpIndex
End Sub


