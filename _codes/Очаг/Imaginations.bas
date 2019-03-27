Attribute VB_Name = "Imaginations"

Sub SquareSetInner(shpObjName As String) 'ВНУТРЕННЯЯ
'Процедура присвоения текстовому полю выделенной фигуры значения площади фигуры
Dim SquareCalc As Integer
Dim ShpObj As Visio.Shape

'---Определяем объектную переменную для активной фигуры
    Set ShpObj = Application.ActivePage.Shapes(shpObjName)

'---Определяем её площадь
    SquareCalc = ShpObj.AreaIU * 0.00064516 'переводим из квадратных дюймов в квадратные метры
    ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub

Sub CloneSectionUniverseNames(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования необходимого набора свойств указанной секции из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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
            
End Sub

Sub CloneSectionUniverseValues(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'---Процедура копирования свойств указанной секции из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Запускаем цикл работы со строками Шейп-листа
        For RowNum = 0 To ShapeFrom.RowCount(SectionIndex) - 1
            
        '---Запускаем цикл работы с ячейками в строке
            For CellNum = 0 To ShapeFrom.RowsCellCount(SectionIndex, RowNum) - 1
                ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Formula = ShapeFrom.CellsSRC(SectionIndex, RowNum, CellNum).Formula
                'MsgBox RowNum & ShapeTo.CellsSRC(SectionIndex, RowNum, CellNum).Name
            Next CellNum
        Next RowNum


End Sub


'---------------------------------Обращение в площадь горения-------------------------------------
Sub ImportAreaInformation() '(Площадь пожара)
'Процедура для импорта свойств объекта донора

'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

''---Проверяем выбран ли какой либо объект
    'If Application.ActiveWindow.Selection.Count < 1 Then
    '    MsgBox "Не выбрана ни одна фигура!", vbInformation
    '    Exit Sub
    'End If
    '
'---Учитываем, что фигура может быть местом
    If Application.ActiveWindow.Selection(1).CellExists("User.visObjectType", 0) Then
        If Application.ActiveWindow.Selection(1).Cells("User.visObjectType") = 104 Then
            PF_GeometryCopy Application.ActiveWindow.Selection(1)
        End If
    End If
'
''---Проверяем, не является ли выбранная фигура уже площадью или другой фигурой с назначенными свойствами
    'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
    '    MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в зону горения", vbInformation
    '    Exit Sub
    'End If
'
''---Проверяем Является ли выбранная фигура площадью
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "Выбранная фигура не имеет площади!", vbInformation
    '    Exit Sub
    'End If


'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters("Площадь прямоугольная").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = Application.Documents("Очаг.vss").Masters("Площадь прямоугольная").Index

'---Создаем необходимый набор пользовательских ячеек для секций User, Prop, Action, Controls
    CloneSectionUniverseNames 240, IDFrom, IDTo  'Action
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 240, IDFrom, IDTo  'Action
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
    CloneSecEvent IDFrom, IDTo
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo
    CreateTextFild IDTo

'---Устанавливаем значение текущего времени из ячейки TheDoc!User.CurrentTime
    ShapeTo.Cells("Prop.SquareTime").Formula = _
        "=DATETIME(" & CStr(Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"

'---Присваиваем номер слоя
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---Открываем окно свойств обращенной фигуры
    On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
    'Application.DoCmd (1312)
    If VfB_NotShowPropertiesWindow = False Then Application.DoCmd (1312) 'В случае если показ окон включен, показываем окно

    SquareSetInner (ShapeTo.Name)

End Sub

Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'Процедура клонирования данных из секции "Line"

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1)
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
Dim RowNum As Integer   ', CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j

End Sub

Sub CloneSecFill(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Fill"
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowFill, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowFill, j).Formula
    Next j

End Sub

Sub CloneSecMiscellanious(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Miscellanious"
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
    Next j

End Sub


Sub CreateTextFild(ShapeToID As Long)
'Процедура текстового поля
'---Объявляем переменные
Dim ShapeTo As Visio.Shape
Dim vsoCharacters As Visio.Characters

'---Присваиваем значения переменным
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)
Set vsoCharacters = ShapeTo.Characters

'MsgBox ShapeTo.Name

'---Создаем новое текстовое поле и присваиваем ему значения
    vsoCharacters.Begin = 0
    vsoCharacters.End = 0
    vsoCharacters.AddCustomFieldU "GUARD(Prop.FireCategorie&Prop.IntenseShowType&Prop.FireDescription)", visFmtNumGenNoUnits
    ShapeTo.CellsSRC(visSectionCharacter, 0, visCharacterLangID).FormulaU = 1033

'---Скрываем текст фигуры от просмотра и доступа
    ShapeTo.CellsSRC(visSectionObject, visRowLock, visLockTextEdit).FormulaU = 1
    ShapeTo.CellsSRC(visSectionObject, visRowMisc, visHideText).FormulaU = True

'---Очищаем переменные
Set ShapeTo = Nothing
Set vsoCharacters = Nothing

End Sub


'---------------------------------Обращение в Шторм-------------------------------------
Sub ImportStormInformation()
'Процедура для импорта свойств обекта донора
'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

''---Проверяем выбран ли какой либо объект
    'If Application.ActiveWindow.Selection.Count < 1 Then
    '    MsgBox "Не выбрана ни одна фигура!", vbInformation
    '    Exit Sub
    'End If
    '
''---Проверяем, не является ли выбранная фигура уже штормом или другой фигурой с назначенными свойствами
    'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
    '    MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в огненный шторм", vbInformation
    '    Exit Sub
    'End If
    '
''---Проверяем Является ли выбранная фигура площадью
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "Выбранная фигура не имеет площади и не может быть обращена в огненный шторм!", vbInformation
    '    Exit Sub
    'End If

'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters("Огненный шторм").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = Application.Documents("Очаг.vss").Masters("Огненный шторм").Index

'---Создаем необходимый набор пользовательских ячеек для секций User и Actions
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 240, IDFrom, IDTo  'Action

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 240, IDFrom, IDTo  'Action

'---Копируем формулы ячеек для указанных секций
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo

'---Присваиваем номер слоя
ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

End Sub


'---------------------------------Обращение в задымленную зону-------------------------------------
Sub ImportFogInformation()
'Процедура для импорта свойств обекта донора

'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
'---Проверяем выбран ли какой либо объект
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "Не выбрана ни одна фигура!", vbInformation
        Exit Sub
    End If

'---Проверяем, не является ли выбранная фигура уже площадью или другой фигурой с назначенными свойствами
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в задымленную зону", vbInformation
        Exit Sub
    End If

'---Проверяем Является ли выбранная фигура площадью
    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
        MsgBox "Выбранная фигура не имеет площади и не может быть обращена!", vbInformation
        Exit Sub
    End If

'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters("Задымление").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = Application.Documents("Очаг.vss").Masters("Задымление").Index

'---Создаем необходимый набор пользовательских ячеек для секций User, Prop, Action, Controls
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo

'---Присваиваем номер слоя
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

Exit Sub
EX:
    SaveLog Err, "ImportFogInformation"
End Sub


'---------------------------------Обращение в зону обрушения-------------------------------------
Sub ImportRushInformation()
'Процедура для импорта свойств обекта донора

'---Объявляем переменные
Dim IDFrom As Long, IDTo As Long
Dim ShapeTo As Visio.Shape, ShapeFrom As Visio.Shape

    On Error GoTo EX
'---Проверяем выбран ли какой либо объект
    If Application.ActiveWindow.Selection.Count < 1 Then
        MsgBox "Не выбрана ни одна фигура!", vbInformation
        Exit Sub
    End If

'---Проверяем, не является ли выбранная фигура уже площадью или другой фигурой с назначенными свойствами
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в зону обрушения", vbInformation
        Exit Sub
    End If

'---Проверяем Является ли выбранная фигура площадью
    'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
    '    MsgBox "Выбранная фигура не имеет площади и не может быть обращена!", vbInformation
    '    Exit Sub
    'End If

'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
    Set ShapeTo = Application.ActiveWindow.Selection(1)
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters("Зона обрушения").Shapes(1)
    IDTo = ShapeTo.ID
    IDFrom = Application.Documents("Очаг.vss").Masters("Зона обрушения").Index

'---Создаем необходимый набор пользовательских ячеек для секций User, Prop, Action, Controls
    CloneSectionUniverseNames 242, IDFrom, IDTo  'User
    CloneSectionUniverseNames 243, IDFrom, IDTo  'Prop

'---Копируем формулы ячеек для указанных секций
    CloneSectionUniverseValues 242, IDFrom, IDTo  'User
    CloneSectionUniverseValues 243, IDFrom, IDTo  'Prop
    CloneSectionLine IDFrom, IDTo
    CloneSecFill IDFrom, IDTo
    CloneSecEvent IDFrom, IDTo
    CloneSecMiscellanious IDFrom, IDTo

'---Указываем текущее время (без ссылки)
    ShapeTo.Cells("Prop.RushTime").Formula = "DateTime(" & ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate) & ")"

'---Присваиваем номер слоя
    ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---Открываем окно свойств обращенной фигуры
    On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ

    SquareSetInner (ShapeTo.Name)
    Application.DoCmd (1312)
    
Exit Sub
EX:
    SaveLog Err, "ImportRushInformation"
End Sub

'--------------------------------Получение номера слоя--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'Функция возвращает номер слоя в текущем документе соответствующего слою в документе мастера
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, layerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.layer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = Application.Documents("Очаг.vss").Masters(ShapeFromID).Shapes(1)

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

End Function



'-------------------------------------Обращение фигуры места в площадь пожара--------------------------------------
Public Function PF_GeometryCopy(ByRef OriginalShp As Visio.Shape) As Visio.Shape
'Прока создает графическую фигуру с копией геометрии исходной фигуры
'Dim OriginalShp As Visio.Shape
Dim ReplicaShape As Visio.Shape
Dim i As Integer
Dim j As Integer
Dim k As Integer

    On Error GoTo EX
'    Set OriginalShp = Application.ActiveWindow.Page.Shapes.ItemFromID(231)
    Set ReplicaShape = Application.ActiveWindow.Page.DrawRectangle(0, 0, 100, 100)
    
    '---Формируем геометрию
    i = 0
    Do While OriginalShp.SectionExists(visSectionFirstComponent + i, 0)
        If i > 0 Then
            ReplicaShape.AddSection (visSectionFirstComponent + i)
            ReplicaShape.AddRow visSectionFirstComponent + i, visRowComponent, visTagComponent
        Else
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
            ReplicaShape.DeleteRow visSectionFirstComponent, 1
        End If
        
        j = 1
        Do While OriginalShp.RowExists(visSectionFirstComponent + i, j, 0)
            ReplicaShape.AddRow visSectionFirstComponent + i, j, OriginalShp.RowType(visSectionFirstComponent + i, j)
            
            k = 0
            Do While OriginalShp.CellsSRCExists(visSectionFirstComponent + i, j, k, 0)
                ReplicaShape.CellsSRC(visSectionFirstComponent + i, j, k).FormulaU = _
                    OriginalShp.CellsSRC(visSectionFirstComponent + i, j, k).FormulaU
                
                k = k + 1
            Loop
            j = j + 1
        Loop
        i = i + 1
    Loop
    
    '---Приравниваем положение и размеры исходной фигуры к данным реплики
        ReplicaShape.Cells("Width").FormulaU = OriginalShp.Cells("Width").FormulaU
        ReplicaShape.Cells("Height").FormulaU = OriginalShp.Cells("Height").FormulaU
        ReplicaShape.Cells("LocPinX").FormulaU = OriginalShp.Cells("LocPinX").FormulaU
        ReplicaShape.Cells("LocPinY").FormulaU = OriginalShp.Cells("LocPinY").FormulaU
        ReplicaShape.Cells("PinX").FormulaU = OriginalShp.Cells("PinX").FormulaU
        ReplicaShape.Cells("PinY").FormulaU = OriginalShp.Cells("PinY").FormulaU
        
Set PF_GeometryCopy = ReplicaShape

Exit Function
EX:
    MsgBox "Возникла непредвиденная ошибка! Если она будет повторяться - обратитесь к разработчику"
    SaveLog Err, "Document_DocumentOpened"
End Function








