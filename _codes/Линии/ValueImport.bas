Attribute VB_Name = "ValueImport"
'------------------------Модуль для процедур импорта значений ячеек-------------------
'------------------------Блок Значений ячеек------------------------------------------
Public Sub HoseResistanceValueImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Сопротивление в соответсвии с материалом и диаметром
'---Объявляем переменные
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения показателя сопротивления пожарного рукава
    Select Case indexPers
        Case Is = 100
            Criteria = "[Материал рукава] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[Диаметр рукавов] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
'            Debug.Print CStr(ValueImportSng("З_Рукава", "Сопротивление", Criteria))
            shp.Cells("Prop.HoseResistance").FormulaU = """" & CStr(ValueImportSng("З_Рукава", "Сопротивление", Criteria)) & """"
    End Select

Set shp = Nothing

End Sub

Public Sub HoseMaxFlowValueImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Пропускная способность в соответсвии с материалом и диаметром
'---Объявляем переменные
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения Типа струи при заданном Виде струи для текущей фигуры
    Select Case indexPers
        Case Is = 100
            Criteria = "[Материал рукава] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[Диаметр рукавов] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.FlowS").FormulaU = str(ValueImportSng("З_Рукава", "Расход", Criteria))
    
    End Select

Set shp = Nothing

End Sub

Public Sub HoseWeightValueImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Масса З_Рукава в соответсвии с материалом и диаметром
'---Объявляем переменные
Dim shp As Visio.Shape
Dim indexPers As Integer
Dim Criteria As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    indexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения Типа струи при заданном Виде струи для текущей фигуры
    Select Case indexPers
        Case Is = 100
            Criteria = "[Материал рукава] = '" & shp.Cells("Prop.HoseMaterial").ResultStr(visUnitsString) & "' And " & _
                "[Диаметр рукавов] = " & shp.Cells("Prop.HoseDiameter").ResultStr(visUnitsString)
            shp.Cells("Prop.HoseWeight").FormulaU = str(ValueImportSng("З_Рукава", "Масса", Criteria))
    
    End Select

Set shp = Nothing

End Sub
