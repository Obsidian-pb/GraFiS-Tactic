Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)

    '---Обновляем общие списки
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

    '---Проверяем для какой фигуры выполняется процедура и обновляем зависимые списки
    Select Case ShpObj.Cells("User.IndexPers")
        Case Is = 100 'Водяной ручной ствол
            ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
            
    End Select
    
    
On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИБКИ
Application.DoCmd (1312)

End Sub

Public Sub InLineListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)

    '---Обновляем общие списки
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИБКИ
Application.DoCmd (1312)

End Sub

'------------------------Блок зависимых списков------------------------------
Public Sub HoseDiametersListImport(ShpIndex As Long)
'Процедура импорта Вариантов стволов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
Select Case indexPers
    Case Is = 100
        Criteria = "[Материал рукава] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.HoseDiameter.Format").FormulaU = ListImport2("З_Рукава", "Диаметр рукавов", Criteria)

End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
If shp.Cells("Prop.HoseDiameter").ResultStr(Visio.visNone) = "" Then
    shp.Cells("Prop.HoseDiameter").FormulaU = "INDEX(0,Prop.HoseDiameter.Format)"
End If

Set shp = Nothing

End Sub


