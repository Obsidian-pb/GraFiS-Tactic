Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)
'---Обновляем общие списки
ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

'---Обновляем список моделей и их ТТХ
ModelsListImport (ShpObj.ID)

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
Select Case IndexPers
    Case Is = 58 'Носимые радиостанции
        Criteria = "Носимая"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("Радиостанции", "Модель", "Тип", Criteria)
    Case Is = 59   '"Автомобильная"
        Criteria = "Автомобильная"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("Радиостанции", "Модель", "Тип", Criteria)
    Case Is = 23 'Стационарная радиостанция
        Criteria = "Стационарная"
        shp.Cells("Prop.Model.Format").FormulaU = ListImport2("Радиостанции", "Модель", "Тип", Criteria)
        
End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
If shp.Cells("Prop.Model").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.Model").FormulaU = "INDEX(0,Prop.Model.Format)"
End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ModelsListImport"
End Sub








