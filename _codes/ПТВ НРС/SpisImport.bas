Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)
'Dim IndexPers As Integer

    '---Обновляем общие списки
'    ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволов", "Модель ствола")
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

    '---Проверяем для какой фигуры выполняется процедура и обновляем зависимые списки
    Select Case ShpObj.Cells("User.IndexPers")
        Case Is = 34 'Водяной ручной ствол
            ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволов", "Модель ствола")
        Case Is = 36 'Лафетный водяной ствол
            ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволовЛ", "Модель ствола")
        Case Is = 35 'Пенный ручной ствол
            ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросПенныхСтволов", "Модель ствола")
            ShpObj.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("Пенообразователи", "Пенообразователь")
        Case Is = 37 'Пенный Лафетный ствол
            ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросПенныхСтволовЛ", "Модель ствола")
            ShpObj.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("Пенообразователи", "Пенообразователь")
        Case Is = 39 'Возимый водяной лафетный ствол
            ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволовЛВ", "Модель ствола")
        Case Is = 76 'Ручной газовый ствол
            ShpObj.Cells("Prop.Gas.Format").FormulaU = ListImport("Газовые составы", "Состав")
        Case Is = 77 'Ручной порошковый ствол
            ShpObj.Cells("Prop.Powder.Format").FormulaU = ListImport("Порошки", "Марка")
        Case Is = 40 'Гидроэлеватор
            ShpObj.Cells("Prop.WEType.Format").FormulaU = ListImport("Гидроэлеваторы", "Модель")
        Case Is = 41 'Пеносмеситель - ПОКА ничего не происходит!!!
            
        Case Is = 88 'Сетка всасывающая
            ShpObj.Cells("Prop.Model.Format").FormulaU = ListImport("Сетки всасывающие", "Модель")
        Case Is = 72 'Колонка пожарная
            ShpObj.Cells("Prop.ColPressure.Format").FormulaU = ListImportInt("Колонки", "Напор")
            
            
    End Select
    
'---Очищаем текущие соединения для фигуры
    ConnectionsRefresh ShpObj
    
    
On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

End Sub

Public Sub UnitsListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (Только подразделения)
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

End Sub

Public Sub WFListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (Сетка всасывающая)
'Dim IndexPers As Integer

    '---Обновляем общие списки
'    ShpObj.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволов", "Модель ствола")
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
    ShpObj.Cells("Prop.WFType.Format").FormulaU = ListImport("Сетки всасывающие", "Модель")
    
    
On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

End Sub

Public Sub StvolModelsListImport(ShpIndex As Long)
'Процедура импорта моделей стволов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
    Select Case IndexPers
        Case Is = 34 'Водяной ручной ствол
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволов", "Модель ствола")
        Case Is = 36 'Лафетный водяной ствол
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволовЛ", "Модель ствола")
        Case Is = 35 'Пенный ствол
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросПенныхСтволов", "Модель ствола")
        Case Is = 37 'Пенный лафетный ствол
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросПенныхСтволовЛ", "Модель ствола")
        Case Is = 39 'Лафетный возимый водяной ствол
            shp.Cells("Prop.StvolType.Format").FormulaU = ListImport("ЗапросВодяныхСтволовЛВ", "Модель ствола")
            
            
            
    End Select
    
    '---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.StvolType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StvolType").FormulaU = "INDEX(0,Prop.StvolType.Format)" '!!! Тут может быть ошибка!!!!
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolModelsListImport"
End Sub

Public Sub WEModelsListImport(ShpIndex As Long)
'Процедура импорта моделей стволов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
    shp.Cells("Prop.WEType.Format").FormulaU = ListImport("ЗапросГЭ", "Модель")
        
'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.WEType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.WEType").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "WEModelsListImport"
End Sub

Public Sub WFModelsListImport(ShpIndex As Long)
'Процедура импорта моделей всасывающих сеток
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
    shp.Cells("Prop.WFType.Format").FormulaU = ListImport("Сетки всасывающие", "Модель")
        
'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.WFType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.WFType").FormulaU = "INDEX(0,Prop.WFType.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "WFModelsListImport"
End Sub

Public Sub StvolFoamCreatorListImport(ShpIndex As Long)
'Процедура импорта моделей стволов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка Модели стволов для текущей фигуры
'Criteria = shp.Cells("Prop.Set").ResultStr(visUnitsString)
Select Case IndexPers
    Case Is = 35 'Пенный ручной ствол
        shp.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("Пенообразователи", "Пенообразователь")
    Case Is = 37 'Пенный лафетный ствол
        shp.Cells("Prop.FoamCreator.Format").FormulaU = ListImport("Пенообразователи", "Пенообразователь")
        
        
        
        
        
End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.StvolType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StvolType").FormulaU = "INDEX(0,Prop.Model.Format)"
    End If
        
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolFoamCreatorListImport"
End Sub




'------------------------Блок зависимых списков------------------------------
Public Sub StvolVariantsListImport(ShpIndex As Long)
'Процедура импорта Вариантов стволов
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
    Case Is = 34
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("ЗапросВодяныхСтволов", "Вариант ствола", Criteria)
    Case Is = 36
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛ", "Вариант ствола", Criteria)
    Case Is = 35 ' Пенный ручной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("ЗапросПенныхСтволов", "Вариант ствола", Criteria)
    Case Is = 37 ' Пенный лафетный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("ЗапросПенныхСтволовЛ", "Вариант ствола", Criteria)
    Case Is = 39 'Лафетный водяной возимый ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' "
        shp.Cells("Prop.Variant.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛВ", "Вариант ствола", Criteria)
        
        
        
        
End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.Variant").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Variant").FormulaU = "INDEX(0,Prop.Variant.Format)"
    End If
    
    Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolVariantsListImport"
End Sub


Public Sub StvolStreamTypesListImport(ShpIndex As Long)
'Процедура импорта списка возможных струй для данного ствола
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
    Case Is = 34
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("ЗапросВодяныхСтволов", "Вид струи", Criteria)
    Case Is = 36
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛ", "Вид струи", Criteria)
    Case Is = 35
        Set shp = Nothing
        Exit Sub
    Case Is = 37
        Set shp = Nothing
        Exit Sub
    Case Is = 39 'Водяной возимый лафетный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.StreamType.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛВ", "Вид струи", Criteria)
        
        
        
        
End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.StreamType").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.StreamType").FormulaU = "INDEX(0,Prop.StreamType.Format)"
    End If
    
    Set shp = Nothing

Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolStreamTypesListImport"
End Sub



Public Sub StvolHeadListImport(ShpIndex As Long)
'Процедура импорта списка возможных напоров для данного вида струи и ствола
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer
Dim Criteria As String

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")

'---Запускаем процедуру получения относительного списка возможных Напоров для текущей фигуры
Select Case IndexPers
    Case Is = 34
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("ЗапросВодяныхСтволов", "Напор", Criteria)
    Case Is = 36
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛ", "Напор", Criteria)
    Case Is = 35 'Пенный ручной ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("ЗапросПенныхСтволов", "Напор", Criteria)
    Case Is = 37 'Пенный лафетный ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("ЗапросПенныхСтволовЛ", "Напор", Criteria)
    Case Is = 39 'Водяной лафетный возимый ствол
        Criteria = "[Модель ствола] = '" & shp.Cells("Prop.StvolType").ResultStr(visUnitsString) & "' And " & _
            "[Вариант ствола] = '" & shp.Cells("Prop.Variant").ResultStr(visUnitsString) & "' And " & _
            "[Вид струи] = '" & shp.Cells("Prop.StreamType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Head.Format").FormulaU = ListImport2("ЗапросВодяныхСтволовЛВ", "Напор", Criteria)
    Case Is = 40 'Гидроэлеватор
        Criteria = "[Модель] = '" & shp.Cells("Prop.WEType").ResultStr(visUnitsString) & "'"
        shp.Cells("Prop.Pressure.Format").FormulaU = ListImport2("ЗапросГЭ", "Напор на входе", Criteria)
        
        
        
End Select

'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.Head").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.Head").FormulaU = "INDEX(0,Prop.Head.Format)"
    End If
    
    Set shp = Nothing
Exit Sub
EX:
    Set shp = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StvolHeadListImport"
End Sub



