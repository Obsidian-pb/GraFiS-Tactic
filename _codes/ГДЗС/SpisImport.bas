Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)

'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Обновляем общие списки
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
        'ShpObj.Cells("Prop.AirDevice.Format").FormulaU = ListImport("ДАСВ", "Модель")
        
        '---Обновляем список моделей и их ТТХ
        '---Проверяем для какой фигуры выполняется процедура и обновляем зависимые списки (только для звеньев)
        Select Case ShpObj.Cells("User.IndexPers")
            Case Is = 46 'ДАСВ
                AirDevicesListImport (ShpObj.ID)
                GetTTH (ShpObj.ID)
            Case Is = 90 'ДАСК
                AirDevicesListImport (ShpObj.ID)
                GetTTH (ShpObj.ID)
        End Select
        
        '---Добавляем ссылку на текущее время страницы                                                                                         'Для всех остальных
            ShpObj.Cells("Prop.FormingTime").Formula = _
                Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
        

    End If



On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

End Sub

Public Sub FogRMKBaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков) для дымососов
'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Обновляем общие списки
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
        
        '---Обновляем список моделей и их ТТХ
        '---Обновляем список моделей и их ТТХ
        FogRMKListImport (ShpObj.ID)
        GetTTH (ShpObj.ID)
        
        '---Добавляем ссылку на текущее время страницы
            ShpObj.Cells("Prop.SetTime").Formula = _
                Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If
    




On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

End Sub

'------------------------Блок зависимых списков------------------------------
Public Sub AirDevicesListImport(ShpIndex As Long)
'Процедура импорта моделей
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
    Select Case IndexPers
        Case Is = 46
            shp.Cells("Prop.AirDevice.Format").FormulaU = ListImport("ДАСВ", "Модель")
        Case Is = 90
            shp.Cells("Prop.AirDevice.Format").FormulaU = ListImport("ДАСК", "Модель")
    End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.AirDevice").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.AirDevice").FormulaU = "INDEX(0,Prop.AirDevice.Format)"
    End If
    
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    SaveLog Err, "AirDevicesListImport", CStr(ShpIndex)
End Sub


Public Sub FogRMKListImport(ShpIndex As Long)
'Процедура импорта моделей дымососов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
    shp.Cells("Prop.FogRMK.Format").FormulaU = ListImport("Дымососы", "Модель")

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.FogRMK").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.FogRMK").FormulaU = "INDEX(0,Prop.FogRMK.Format)"
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    SaveLog Err, "FogRMKListImport", CStr(ShpIndex)
End Sub
