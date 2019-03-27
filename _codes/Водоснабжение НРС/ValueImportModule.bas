Attribute VB_Name = "ValueImportModule"
'------------------------Модуль для процедур импорта значений ячеек-------------------
'------------------------Блок Значений ячеек------------------------------------------
Public Sub ProductionImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Водоотдача в соответсвии с видом водовода и напором
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения Типа струи при заданном Виде струи для текущей фигуры
    Select Case IndexPers
        Case Is = 50 'Пожарный гидрант
            Criteria = "[Вид водовода] = '" & shp.Cells("Prop.PipeType").ResultStr(visUnitsString) & "' And " & _
                "[Диаметр водовода] = " & shp.Cells("Prop.PipeDiameter").ResultStr(visUnitsString) & " And " & _
                "[Напор в сети] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
            shp.Cells("Prop.Production").FormulaForceU = "Guard(" & ValueImportSng("ЗапросВодоотдачи", "Водоотдача", Criteria) & ")"
    
    End Select

Set shp = Nothing
Exit Sub
EX:
    SaveLog Err, "ProductionImport"
End Sub


