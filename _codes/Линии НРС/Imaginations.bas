Attribute VB_Name = "Imaginations"

Sub LenightSet(ShpObj As Visio.Shape, PRO As Integer)
'Внешняя процедура присвоения текстовому полю выделенной фигуры значения длины фигуры
    Dim LenightCalc As Integer
    
'    LenightCalc = ActivePage.Shapes.ItemFromID(PRO).LengthIU / 39.37 'Переводим из дюймов в метры
'    Application.ActiveWindow.Page.Shapes.ItemFromID(PRO).Cells("User.LineLenight").FormulaForceU = LenightCalc
    LenightCalc = ShpObj.LengthIU / 39.37 'Переводим из дюймов в метры
    ShpObj.Cells("User.LineLenight").FormulaForceU = LenightCalc
    
End Sub

Private Sub LenightSetInner(PRO As String)
'Внутренняя процедура присвоения текстовому полю выделенной фигуры значения длины фигуры
    Dim LenightCalc As Integer
    
    LenightCalc = ActivePage.Shapes(PRO).LengthIU / 39.37 'Переводим из дюймов в метры
    Application.ActiveWindow.Page.Shapes(PRO).Cells("User.LineLenight").FormulaForceU = LenightCalc

End Sub

Sub CloneSectionUniverseNames(ByVal SectionIndex As Integer, ByVal ShapeFromID As Long, ByVal ShapeToID As Long)
'Процедура копирования необходимого набора свойств указанной секции из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

On Error GoTo EX

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Проверяем наличие секции с указанным SectionIndex и в случае отсутствия создаем её
    If (ShapeTo.SectionExists(SectionIndex, 0) = 0) And Not (ShapeFrom.SectionExists(SectionIndex, 0) = 0) Then
        'MsgBox "Создаем новую секцию"
        ShapeTo.AddSection (SectionIndex)
    End If

'On Error Resume Next
'---Запускаем цикл работы со строками Шейп-листа
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            'If (ShapeTo.RowExists(SectionIndex, RowNum, 0) = 0) And Not (ShapeFrom.RowExists(SectionIndex, RowNum, 0) = 0) Then
                'MsgBox "create"
                ShapeTo.AddRow SectionIndex, RowNum, 0
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).RowNameU = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).RowName
            'End If
        Next RowNum
        
    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionUniverseNames"
    
End Sub

Sub CloneSectionScratchNames(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования необходимого набора свойств секции Scratch из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

On Error GoTo EX

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Проверяем наличие секции с указанным SectionIndex и в случае отсутствия создаем её
    If (ShapeTo.SectionExists(visSectionScratch, 0) = 0) And Not (ShapeFrom.SectionExists(visSectionScratch, 0) = 0) Then
        ShapeTo.AddSection (visSectionScratch)
    End If

'---Запускаем цикл работы со строками Шейп-листа
    For RowNum = 0 To ShapeFrom.RowCount(visSectionScratch) - 1
            ShapeTo.AddRow visSectionScratch, RowNum, 0
    Next RowNum
            
    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionScratchNames"
    
End Sub

Sub CloneSectionUniverseValues(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'---Процедура копирования свойств указанной секции из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

On Error GoTo 10

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Запускаем цикл работы со строками Шейп-листа
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            
        '---Запускаем цикл работы с ячейками в строке
            For CellNum = 0 To ShapeFrom.RowsCellCount(SectionIndex, RowNum) - 1
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Formula = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).Formula
                'MsgBox RowNum & ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Name
            Next CellNum
        Next RowNum

    Exit Sub
10:
'    Debug.Print Err.Description
    SaveLog Err, "CloneSectionUniverseValues"

End Sub

Sub CloneSectionScratchValues(ShapeFromID As Long, ShapeToID As Long)
'---Процедура копирования свойств секции Scratch из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

On Error GoTo EX

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
    Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Запускаем цикл работы со строками Шейп-листа
        For RowNum = 0 To ShapeFrom.RowCount(visSectionScratch) - 1
            
        '---Запускаем цикл работы с ячейками в строке
            For CellNum = 0 To ShapeFrom.RowsCellCount(visSectionScratch, RowNum) - 1
                ShapeTo.CellsSRC(visSectionScratch, RowNum, CellNum).Formula = _
                    ShapeFrom.CellsSRC(visSectionScratch, RowNum, CellNum).Formula
            Next CellNum
        Next RowNum

    Exit Sub
EX:
'    Debug.Print Err.Description
    SaveLog Err, "Document_DocumentOpened"

End Sub

Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'Процедура клонирования данных из секции "Line"

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Создаем необхоимое количество строк
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowLine, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowLine, j).Formula
    Next j

End Sub

Sub CloneSecEvent(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Event"
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j

End Sub

Sub CloneSecMiscellanious(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Miscellanious" - копируется только Comment
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
'Dim RowNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Присваиваем полю Comment новой фигуры значение из фигура эталона.
    ShapeTo.CellsSRC(visSectionObject, visRowMisc, visComment).Formula = _
        ShapeFrom.CellsSRC(visSectionObject, visRowMisc, visComment).Formula

'---Очищаем объектные переменные
Set ShapeFrom = Nothing
Set ShapeTo = Nothing

End Sub


Sub ImportHoseInformation()
'Процедура для импорта свойств обекта донора (Рукавная линия)

'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
    '---Проверяем выбран ли какой либо объект
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "Не выбрана ни одна фигура!", vbInformation
        Exit Sub
    End If
    
    '---Проверяем, не является ли выбранная фигура уже рукавом или другой фигурой с назначенными свойствами
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в рукавную линию", vbInformation
        Exit Sub
    End If
    
    '---Проверяем Является ли выбранная фигура линией
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        MsgBox "Выбранная фигура не является линией!", vbInformation
        Exit Sub
    End If

    '---Проверяем имеется ли у данной страницы ячейка Аспекта User.GFS_Aspect, если нет, то создаем ее
    If Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) = False Then
        If Application.ActivePage.PageSheet.SectionExists(visSectionUser, 0) = False Then
            Application.ActivePage.PageSheet.AddSection (visSectionUser)
        End If
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", visTagDefault
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").Formula = 1
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect.Prompt").FormulaU = """Аспект"""
    End If

'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("Рукав - скатка").Shapes(1)
    IDTo = ShapeTo.ID   'Application.ActivePage.Shapes("Sheet.2").ID
    'IDFrom = ShapeFrom.Index
    IDFrom = ThisDocument.Masters("Рукав - скатка").Index

'---Создаем необходимый набор пользовательских ячеек для секций User, Prop, Action, Controls
    CloneSectionUniverseNames 240, IDFrom, IDTo
    CloneSectionUniverseNames 242, IDFrom, IDTo
    CloneSectionUniverseNames 243, IDFrom, IDTo
    CloneSectionScratchNames IDFrom, IDTo

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 240, IDFrom, IDTo
    CloneSectionUniverseValues 242, IDFrom, IDTo
    CloneSectionUniverseValues 243, IDFrom, IDTo
    CloneSectionScratchValues IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSectionLine IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---Присваиваем номер слоя
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---Осуществляем соединение рукавной линии
    ReconnectHose ShapeTo

''---Открываем окно свойств обращенной фигуры
'On Error Resume Next
'Application.DoCmd (1312)

    LenightSetInner (ShapeTo.Name)

'---Очищаем объектные переменные
    Set ShapeTo = Nothing
    Set ShapeFrom = Nothing
    
Exit Sub
EX:
    '---Очищаем объектные переменные
    Set ShapeTo = Nothing
    Set ShapeFrom = Nothing
    SaveLog Err, "ImportHoseInformation"
End Sub


Sub ImportVHoseInformation()
'Процедура для импорта свойств обекта донора (Всасывающий рукав)

'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
    '---Проверяем выбран ли какой либо объект
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "Не выбрана ни одна фигура!", vbInformation
        Exit Sub
    End If
    
    '---Проверяем, не является ли выбранная фигура уже рукавом или другой фигурой с назначенными свойствами
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в рукавную линию", vbInformation
        Exit Sub
    End If
    
    '---Проверяем Является ли выбранная фигура линией
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        MsgBox "Выбранная фигура не является линией!", vbInformation
        Exit Sub
    End If


'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = ThisDocument.Masters("Всасывающая линия").Shapes(1)
    IDTo = ShapeTo.ID
    'IDFrom = ShapeFrom.Index  'Application.Documents("Очаг.vss").Masters("Площадь прямоугольная").Index
    IDFrom = ThisDocument.Masters("Всасывающая линия").Index

'---Создаем необходимый набор пользовательских ячеек для секций User, Prop, Action, Controls
    CloneSectionUniverseNames 240, IDFrom, IDTo
    CloneSectionUniverseNames 242, IDFrom, IDTo
    CloneSectionUniverseNames 243, IDFrom, IDTo
    CloneSectionScratchNames IDFrom, IDTo

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 240, IDFrom, IDTo
    CloneSectionUniverseValues 242, IDFrom, IDTo
    CloneSectionUniverseValues 243, IDFrom, IDTo
    CloneSectionScratchValues IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSectionLine IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---Присваиваем номер слоя
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---Осуществляем соединение рукавной линии
    ReconnectHose ShapeTo

'---Открываем окно свойств обращенной фигуры
    'Application.DoCmd (1312)

    LenightSetInner (ShapeTo.Name)
    
Exit Sub
EX:
    SaveLog Err, "ImportVHoseInformation"
End Sub


'--------------------------------Получение номера слоя--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'Функция возвращает номер слоя в текущем документе соответствующего слою в документе мастера
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, layerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.layer

    On erro GoTo EX
'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = ThisDocument.Masters(ShapeFromID).Shapes(1)
'    MsgBox ThisDocument.Masters(ShapeFromID)
    
'---Получаем название слоя соответственно номеру в исходном документе
    layerName = ShapeFrom.layer(1).Name

'---Проверяем есть ли в текущем документе слой с таким именем
    For i = 1 To Application.ActivePage.Layers.Count
        If Application.ActivePage.Layers(i).Name = layerName Then
            Flag = True
        End If
    Next i

'---В соответствии с полученным названием определяем номер слоя в текущем документе
    If Flag = True Then
        LayerNumber = Application.ActivePage.Layers(layerName).Index
    Else
    '---Создаем новый слой с именем слоя к которому принадлежит исходная фигура
        Set vsoLayer = Application.ActiveWindow.Page.Layers.Add(layerName)
        vsoLayer.NameU = layerName
        vsoLayer.CellsC(visLayerColor).FormulaU = "255"
        vsoLayer.CellsC(visLayerStatus).FormulaU = "0"
        vsoLayer.CellsC(visLayerVisible).FormulaU = "1"
        vsoLayer.CellsC(visLayerPrint).FormulaU = "1"
        vsoLayer.CellsC(visLayerActive).FormulaU = "0"
        vsoLayer.CellsC(visLayerLock).FormulaU = "0"
        vsoLayer.CellsC(visLayerSnap).FormulaU = "1"
        vsoLayer.CellsC(visLayerGlue).FormulaU = "1"
        vsoLayer.CellsC(visLayerColorTrans).FormulaU = "0%"
    '---Присваиваем номер нового слоя
        LayerNumber = Application.ActivePage.Layers(layerName).Index
    End If
        
LayerImport = Chr(34) & LayerNumber - 1 & Chr(34)
Exit Function
EX:
    SaveLog Err, "LayerImport"
End Function


'-----------------------Соединение рукавных линий и ПТВ----------------------------------------------------------
Private Sub ReconnectHose(ByRef ShpObj As Visio.Shape)
'Процедура запускает реконнект соединений для данной фигуры рукавной линии
Dim C_ConnectionsTrace As c_HoseConnector
Dim vO_Conn As Visio.Connect

    '---Создаем экземпляр класса c_HoseConnector для осуществления соединений фигур
    Set C_ConnectionsTrace = New c_HoseConnector
    '---Для всех соединений иимеющихся у рукавной линии запускаем процедуру соединения
    For Each vO_Conn In ShpObj.Connects
        C_ConnectionsTrace.Ps_ConnectionAdd vO_Conn
    Next vO_Conn

'---Очищаем объектные переменные
Set C_ConnectionsTraceLoc = Nothing
End Sub



'-----------------------Процедуры обращения линий----------------------------------------------------------
'Public Sub MakeHoseLine()
''Метод обращения в рукавную линию
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
''    On Error GoTo Tail
'
''---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---Идентифицируем активную фигуру
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---Получаем списки для фигуры
'    '---Запускаем процедуру получения списка Подразделений
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
'    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
'    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
'    HoseDiametersListImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
'    HoseResistanceValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
'    HoseMaxFlowValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
'    HoseWeightValueImport (ShpInd)
'
''---Устанавливаем значение текущего времени бкз ссылки
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---Отмечаем кнопку
''    Ctrl.State = False
'
''---Открываем окно свойств обращенной фигуры
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---Выходим из процедуры обработки
'    Application.EventsEnabled = True
'End Sub

Public Sub MakeVHoseLine()
'Метод обращения во всасывающую рукавную линию
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail

'---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
    Application.EventsEnabled = False
    ImportVHoseInformation
    Application.EventsEnabled = True
    
'---Идентифицируем активную фигуру
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---Получаем списки для фигуры
    '---Запускаем процедуру получения списка Подразделений
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

'---Устанавливаем значение текущего времени бкз ссылки
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
    
'---Открываем окно свойств обращенной фигуры
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
    Application.EventsEnabled = True
End Sub

Public Sub MakeNapVsasHoseLine()
'Метод обращения в напорно-всасывающую рукавную линию
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail

'---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
    Application.EventsEnabled = False
    ImportVHoseInformation
    Application.EventsEnabled = True
    
'---Идентифицируем активную фигуру
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---Получаем списки для фигуры
    '---Запускаем процедуру получения списка Подразделений
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")

'---Устнавливаем значения для напорно-всасывающей линии
    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(1,Prop.LineType.Format)"
    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(0,Prop.HoseDiameter.Format)"

'---Устанавливаем значение текущего времени без ссылки
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
    
'---Открываем окно свойств обращенной фигуры
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
    Application.EventsEnabled = True
End Sub

Public Sub MakeHoseLine(ByVal hoseDiameterIndex As Integer, ByVal lineType As Byte)
'Метод обращения в рукавную линию с заданными параметрами
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer
Dim diameter As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail

'---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
    Application.EventsEnabled = False
    ImportHoseInformation
    Application.EventsEnabled = True
    
'---Идентифицируем активную фигуру
    Set ShpObj = Application.ActiveWindow.Selection(1)
    ShpInd = ShpObj.ID
    
'---Получаем списки для фигуры
    '---Запускаем процедуру получения списка Подразделений
    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
    HoseDiametersListImport (ShpInd)
    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
    HoseResistanceValueImport (ShpInd)
    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
    HoseMaxFlowValueImport (ShpInd)
    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
    HoseWeightValueImport (ShpInd)
        
'---Устнавливаем значения для магистральной линии
    diameter = Index(hoseDiameterIndex, ShpObj.Cells("Prop.HoseDiameter.Format").Formula, ";")
    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(" & diameter & ",Prop.HoseDiameter.Format)"
    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(" & lineType & ",Prop.LineType.Format)"
        
'---Устанавливаем значение текущего времени без ссылки
    ShpObj.Cells("Prop.LineTime").FormulaU = _
        "DATETIME(" & str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
        
'---Открываем окно свойств обращенной фигуры
    On Error Resume Next
    Application.DoCmd (1312)
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
    Application.EventsEnabled = True
End Sub

'Public Sub MakeShortHoseLine()
''Метод обращения в Напорную линию из полурукавчиков рукавную линию
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
'    On Error GoTo Tail
'
''---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---Идентифицируем активную фигуру
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---Получаем списки для фигуры
'    '---Запускаем процедуру получения списка Подразделений
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
'    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
'    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
'    HoseDiametersListImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
'    HoseResistanceValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
'    HoseMaxFlowValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
'    HoseWeightValueImport (ShpInd)
'
''---Устнавливаем значения для магистральной линии
'    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(2,Prop.HoseDiameter.Format)"
'    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(2,Prop.LineType.Format)"
'
''---Устанавливаем значение текущего времени без ссылки
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---Открываем окно свойств обращенной фигуры
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---Выходим из процедуры обработки
'    Application.EventsEnabled = True
'End Sub

