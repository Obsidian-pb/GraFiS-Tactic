Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)
'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Обновляем общие списки
        ShpObj.Cells("Prop.Set.Format").FormulaU = ListImport("Наборы", "Набор")
        ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
        
        '---Обновляем список моделей и их ТТХ
        ModelsListImport (ShpObj.ID)
        GetTTH (ShpObj.ID)
        
        '---Добавляем ссылку на текущее время страницы
        ShpObj.Cells("Prop.ArrivalTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If


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

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 73 ' Машины на гусеничном ходу
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Тип] = 'Машина на гусеничном ходу'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Гусеничные машины", "Модель", Criteria)
        Case Is = 74 ' Танки
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Тип] = 'Танк'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Гусеничные машины", "Модель", Criteria)
        Case Is = 30 'Корабли
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Класс] = 'Море'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Суда", "Проект", Criteria)
        Case Is = 31 'Катер
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Класс] = 'Река'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Суда", "Проект", Criteria)
        Case Is = 24 'Поезда
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Поезда", "Категория", Criteria)
        Case Is = 28 'Мотопомпы
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Мотопомпы", "Модель", Criteria)
        Case Is = 25 'Самолеты
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Тип] = 'Обычный'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Самолеты", "Модель", Criteria)
        Case Is = 26 'Самолеты-Амфибия
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "' And " & _
                "[Тип] = 'Амфибия'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Самолеты", "Модель", Criteria)
        Case Is = 27 'Вертолеты
            Criteria = "[Набор] = '" & shp.Cells("Prop.Set").ResultStr(visUnitsString) & "'"
            shp.Cells("Prop.Model.Format").FormulaU = ListImport3("З_Вертолеты", "Модель", Criteria)
    
    End Select

'---В случае, если значение поля или формата для нового списка равно "", переводим фокус в ячейке на 0-е положение.
If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.Model.Format").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ModelsListImport"
End Sub








