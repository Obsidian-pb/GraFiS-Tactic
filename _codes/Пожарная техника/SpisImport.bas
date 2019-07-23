Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------


Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)
'Dim isExternalData As Boolean
'Dim dataForSave(2) As String
'
'    On Error GoTo EX
'
''---Проверяем, не импортируются ли данные извен (из окна"Внешние данные"
'    isExternalData = WindowCheck("Внешние данные")
'
'    '---Сохраняем данные фигуры
'    If isExternalData Then
'        Exit Sub
'        dataForSave(0) = ShpObj.Cells("Prop.Unit").ResultStr(visUnitsString)
'        dataForSave(1) = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
'    End If

'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        If Not ShpObj.CellExists("User.visDGDefaultPos", 0) Then
            '---Обновляем общие списки
            ShpObj.Cells("Prop.Set.Format").FormulaU = ListImport("Наборы", "Набор")
            ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
            
            '---Обновляем список моделей и их ТТХ
    '            If Not isExternalData Then
'                If Not shp.CellExists("User.visDGDefaultPos", 0) Then
                    ModelsListImport (ShpObj.ID)
                    GetTTH (ShpObj.ID)
                '---Добавляем ссылку на текущее время страницы
                    ShpObj.Cells("Prop.ArrivalTime").Formula = _
                        Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
'                End If
        End If
    End If

'---Очищаем значения строки подключенного лафетного ствола (для предотвращения излишнего расхода в автомобиле)
'    ps_WaterClear ShpObj

'---Восстанавливаем сохраненные данные фигуры
'    If isExternalData Then
'        Application.EventsEnabled = False
'        ShpObj.Cells("Prop.Unit").Formula = """" & dataForSave(0) & """"
'        ShpObj.Cells("Prop.Model").FormulaU = """" & dataForSave(1) & """"
'        Application.EventsEnabled = True
'    End If

'---Очищаем значения подключенных рукавных линий
    ConnectionsRefresh ShpObj

On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

Exit Sub
EX:
    Application.EventsEnabled = True
    SaveLog Err, "BaseListsRefresh", "Пожарная техника"
End Sub
Public Sub BaseListsRefresh2(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков) ТОЛЬКО ДЛЯ ОБЩЕЙ ФИГУРЫ
'---Обновляем общие списки
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
    
    On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
    Application.DoCmd (1312)

End Sub



'------------------------Блок зависимых списков------------------------------
Public Sub ModelsListImport(ShpIndex As Long)
'Процедура импорта моделей
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo Tail

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
    Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 1
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_Автоцистерны", "Модель", "Набор", Criteria)
        Case Is = 2
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АНР", "Модель", "Набор", Criteria)
        Case Is = 3
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АЛ", "Модель", "Набор", Criteria)
        Case Is = 4
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АКП", "Модель", "Набор", Criteria)
        Case Is = 5
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АСО", "Модель", "Набор", Criteria)
        Case Is = 6
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АТ", "Модель", "Набор", Criteria)
        Case Is = 7
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АД", "Модель", "Набор", Criteria)
        Case Is = 8
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_ПНС", "Модель", "Набор", Criteria)
        Case Is = 9
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АА", "Модель", "Набор", Criteria)
        Case Is = 10
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АВ", "Модель", "Набор", Criteria)
        Case Is = 11
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АКТ", "Модель", "Набор", Criteria)
        Case Is = 12
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АП", "Модель", "Набор", Criteria)
        Case Is = 13
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АГВТ", "Модель", "Набор", Criteria)
        Case Is = 14
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АГТ", "Модель", "Набор", Criteria)
        Case Is = 15
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АГДЗС", "Модель", "Набор", Criteria)
        Case Is = 16
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_ПКС", "Модель", "Набор", Criteria)
        Case Is = 17
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_ЛБ", "Модель", "Набор", Criteria)
        Case Is = 18
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АСА", "Модель", "Набор", Criteria)
        Case Is = 19
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АШ", "Модель", "Набор", Criteria)
        Case Is = 20
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АР", "Модель", "Набор", Criteria)
        Case Is = 161 'АЦЛ
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АЦЛ", "Модель", "Набор", Criteria)
        Case Is = 162 'АЦКП
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АЦКП", "Модель", "Набор", Criteria)
        Case Is = 163 'АПП
            shp.Cells("Prop.Model.Format").FormulaU = ListImport2("З_АПП", "Модель", "Набор", Criteria)
            
            
    End Select
    
'---В случае, если значение поля или формата для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.Model.Format").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
    
    Set shp = Nothing
    
Exit Sub
Tail:
    SaveLog Err, "ModelsListImport"
End Sub



