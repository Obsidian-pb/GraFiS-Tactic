Attribute VB_Name = "ValueImportModule"
'------------------------Модуль для процедур импорта значений ячеек-------------------
'------------------------Блок Значений ячеек------------------------------------------
Public Sub StvolStreamValueImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Тип струи в соответсвии с видом струи
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
    Case Is = 34
        Criteria = "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("Струи", "Тип струи", Criteria)
    Case Is = 36
        Criteria = "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("Струи", "Тип струи", Criteria)
    Case Is = 39
        Criteria = "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Stream").FormulaU = ValueImportStr("Струи", "Тип струи", Criteria)
        
        
        
        
End Select

Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolStreamValueImport"
End Sub

Public Sub StvolDiameterInImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Тип струи в соответсвии с видом струи
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
    Case Is = 34
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("МоделиСтволов", "Условный проход", Criteria)
    Case Is = 36
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("МоделиСтволов", "Условный проход", Criteria)
    Case Is = 35 'Пенный ручной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("МоделиСтволов", "Условный проход", Criteria)
    Case Is = 37 'Пенный лафетный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("МоделиСтволов", "Условный проход", Criteria)
    Case Is = 39 'Лафетный возимый ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.DiameterIn").FormulaU = ValueImportSng("МоделиСтволов", "Условный проход", Criteria)
        
        
End Select

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolDiameterInImport"
End Sub

Public Sub StvolProductionImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Напор в соответсвии с прочими данными
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
    Case Is = 34   ' Ручной водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[Напор] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("ЗапросВодяныхСтволов", "Расход", Criteria))
    Case Is = 36   ' Лафетный водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[Напор] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("ЗапросВодяныхСтволовЛ", "Расход", Criteria))
    Case Is = 35   ' Ручной пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Напор] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("ЗапросПенныхСтволов", "Расход", Criteria))
    Case Is = 37   ' Лафетный пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Напор] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("ЗапросПенныхСтволовЛ", "Расход", Criteria))
    Case Is = 39   ' Лафетный возимый водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "' And " & _
            "[Напор] = " & shp.Cells("Prop.Head").ResultStr(visUnitsString)
        shp.Cells("Prop.PodOut").FormulaU = str(ValueImportSng("ЗапросВодяныхСтволовЛВ", "Расход", Criteria))
    Case Is = 40   ' Гидроэлеватор
        'Ниже - заменить на коэффициенты эжекции и подпора
'        Criteria = "[Модель] = '" & shp.Cells("Prop.WEType").ResultStr(visUnitsString) & "'  And " & _
'            "[Напор на входе] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
'        shp.Cells("Prop.PodOut").FormulaU = Str(ValueImportSng("ЗапросГЭ", "Производительность", Criteria))
'        shp.Cells("Prop.PressureOut").FormulaU = Str(ValueImportSng("ЗапросГЭ", "Производительность", Criteria))
    Case Is = 88   ' Всасывающая сетка
        Criteria = "[Модель] = '" & shp.Cells("Prop.WFType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.PodIn").FormulaU = str(ValueImportSng("Сетки всасывающие", "Производительность", Criteria))



End Select


Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolProductionImport"
End Sub

Public Sub StvolProvKoeffImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке User.ProvKoeff в соответсвии с прочими данными значение Проводимости насадка
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
    Case Is = 34   ' Ручной водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("User.ProvKoeff").FormulaU = str(ValueImportSngStr("ЗапросВодяныхСтволов", "Проводимость", Criteria))
    Case Is = 36   ' Лафетный водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("User.ProvKoeff").FormulaU = str(ValueImportSngStr("ЗапросВодяныхСтволовЛ", "Проводимость", Criteria))
    Case Is = 35   ' Ручной пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("User.ProvKoeff").FormulaU = str(ValueImportSngStr("ЗапросПенныхСтволов", "Проводимость", Criteria))
    Case Is = 37   ' Лафетный пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("User.ProvKoeff").FormulaU = str(ValueImportSngStr("ЗапросПенныхСтволовЛ", "Проводимость", Criteria))
    Case Is = 39   ' Лафетный возимый водяной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("User.ProvKoeff").FormulaU = str(ValueImportSngStr("ЗапросВодяныхСтволовЛВ", "Проводимость", Criteria))
    Case Is = 40   ' Гидроэлеватор
        'Ниже - заменить на коэффициенты эжекции и подпора
'        Criteria = "[Модель] = '" & shp.Cells("Prop.WEType").ResultStr(visUnitsString) & "'  And " & _
'            "[Напор на входе] = " & shp.Cells("Prop.Pressure").ResultStr(visUnitsString)
'        shp.Cells("Prop.PodOut").FormulaU = Str(ValueImportSng("ЗапросГЭ", "Производительность", Criteria))
'        shp.Cells("Prop.PressureOut").FormulaU = Str(ValueImportSng("ЗапросГЭ", "Производительность", Criteria))
    Case Is = 88   ' Всасывающая сетка
        Criteria = "[Модель] = '" & shp.Cells("Prop.WFType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.PodIn").FormulaU = str(ValueImportSng("Сетки всасывающие", "Производительность", Criteria))
        
        
        
End Select


Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolProvKoeffImport"
End Sub

Public Sub StvolWFLinkImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке WFLink (ссылка на страничку сайта wiki-fire.org)
'в соответствии с моделью ствола
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
    
'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
'    IndexPers = shp.Cells("User.IndexPers")
    
'---Формируем запрос к БД и получаем значение ссылки на wiki-fire.org
    Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "'"
    shp.Cells("Prop.WFLink").FormulaU = ValueImportStr("МоделиСтволов", "Ссылка WF", Criteria)
   
Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolWFLinkImport"
End Sub

Public Sub StvolWFLinkFree(ShpIndex As Long)
'Прока устанавливает в качестве ссылки пустое значение
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX
    
'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    
'---Устанавливаем пустую ссылку, чтобы формула получения ссылки использовала модель ствола по-умолчанию
    shp.Cells("Prop.WFLink").FormulaU = ""
   
Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolWFLinkImport"
End Sub

Public Sub StvolRFImport(ShpIndex As Long)
'Процедура возвращающая и присваивающая ячейке Кратность в соответсвии с моделью ствола
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения Кратности при заданной модели ствола для текущей фигуры
Select Case IndexPers

    Case Is = 35   ' Ручной пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.FoamRF").FormulaU = str(ValueImportSng("ЗапросПенныхСтволов", "Кратность", Criteria))
    Case Is = 37   ' Лафетный пенный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.FoamRF").FormulaU = str(ValueImportSng("ЗапросПенныхСтволовЛ", "Кратность", Criteria))
        
        
        
        
End Select

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolRFImport"
End Sub


Public Sub ColFlowMaxImport(ShpIndex As Long)
'Процедура присваивающая значение ячейке Проводимость в соответсвии с напором
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
'    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения Кратности при заданной модели ствола для текущей фигуры
'---Определяем критерий отбора - напор перед колонкой
    Criteria = "[Напор] = " & shp.Cells("Prop.ColPressure").ResultStr(visUnitsString)
'---Проверяем какие патрубки у колонки и в соответствии с этим получаем значение из БД
    If shp.Cells("Prop.Patr").ResultStr(visUnitsString) = "77" Then
        shp.Cells("Prop.FlowMax").FormulaU = str(ValueImportSngStr("Колонки", "Расход 77", Criteria))
    ElseIf shp.Cells("Prop.Patr").ResultStr(visUnitsString) = "66" Then
        shp.Cells("Prop.FlowMax").FormulaU = str(ValueImportSngStr("Колонки", "Расход 66", Criteria))
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ColFlowMaxImport"
End Sub

Public Sub FoamDiffKoeffImport(ShpObj As Visio.Shape)
'Процедура присваивающая значение ячейке Коэффициент разности напорв для рукавных вставок (и пеносмесителей)
'---Объявляем переменные
Dim Criteria As String
Dim foamPercentage As String
Dim vstDiameter As Integer

    On Error GoTo EX

'---Запускаем процедуру получения коэффициента при заданных параметрах вставки (пеносмесителя)
    foamPercentage = ShpObj.Cells("Prop.FoamPercentage").ResultStr(visUnitsString)
    vstDiameter = ShpObj.Cells("Prop.FoamInDiameter").Result(visNumber)
    
'---Определяем критерий отбора - напор перед колонкой
    Criteria = "[Концентрация вставки] = " & foamPercentage
'---Проверяем какого диаметра пенная вставка
    If vstDiameter = "10" Then
        ShpObj.Cells("User.DiffKoeff").Formula = str(ValueImportSngStr("КонцентрацияВставки", "Коэффициент разности 10", Criteria))
    End If
    If vstDiameter = "25" Then
        ShpObj.Cells("User.DiffKoeff").Formula = str(ValueImportSngStr("КонцентрацияВставки", "Коэффициент разности 25", Criteria))
    End If

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "FoamDiffKoeffImport"
End Sub

Public Sub NozzleProvKoeffImport(ShpObj As Visio.Shape)
'Процедура присваивающая значение ячейке Проводимость насадка - User.ProvKoeff
'---Объявляем переменные
Dim Criteria As String
Dim nozzleDiameter As String

    On Error GoTo EX

'---Запускаем процедуру получения коэффициента при заданных параметрах насадка (диаметр)
    nozzleDiameter = ShpObj.Cells("User.NozzleDiameter").ResultStr(visUnitsString)
    
'---Определяем критерий отбора - напор перед колонкой
    Criteria = "[Диаметр насадка] = " & nozzleDiameter
'---Импортируем значение
    ShpObj.Cells("User.ProvKoeff").Formula = str(ValueImportSngStr("КонстантыНасадков", "Проводимость", Criteria))

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "NozzleProvKoeffImport"
End Sub

Public Sub NozzleProvKoeffImport2(ShpObj As Visio.Shape)
'Процедура присваивающая значение ячейке Проводимость насадка - User.ProvKoeff
'ДЛЯ СТВОЛОВ ПО МОДЕЛИ
'---Объявляем переменные
Dim Criteria As String
Dim nozzleDiameter As String

    On Error GoTo EX

'---Запускаем процедуру получения коэффициента при заданных параметрах насадка (диаметр)
    nozzleDiameter = ShpObj.Cells("User.NozzleDiameter").ResultStr(visUnitsString)
    
'---Определяем критерий отбора - напор перед колонкой
    Criteria = "[Диаметр насадка] = " & nozzleDiameter
'---Импортируем значение
    ShpObj.Cells("User.ProvKoeff").Formula = str(ValueImportSngStr("КонстантыНасадков", "Проводимость", Criteria))

Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "NozzleProvKoeffImport"
End Sub
