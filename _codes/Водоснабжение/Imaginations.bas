Attribute VB_Name = "Imaginations"

'---------------------------------Обращение в открытый водоисточник-------------------------------------
Public Sub ImportOpenWaterInformation()
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
''---Проверяем, не является ли выбранная фигура уже площадью или другой фигурой с назначенными свойствами
'If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
'    MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в естественный водоем", vbInformation
'    Exit Sub
'End If
'
''---Проверяем Является ли выбранная фигура площадью
'If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
'    MsgBox "Выбранная фигура не имеет площади и не может быть обращена в естественный водоем!", vbInformation
'    Exit Sub
'End If


'---Присваиваем переменным индексы Фигур(ShapeFrom и ShpeTo)
Set ShapeTo = Application.ActiveWindow.Selection(1)
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters("Открытый водоисточник").Shapes(1)
IDTo = ShapeTo.ID
IDFrom = Application.Documents("Водоснабжение.vss").Masters("Открытый водоисточник").Index

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
CloneSecMiscellanious IDFrom, IDTo
CloneSecFill IDFrom, IDTo
'CreateTextFild IDTo

'---Присваиваем номер слоя
ShapeTo.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = LayerImport(IDFrom, IDTo)

'---Открываем окно свойств обращенной фигуры
On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИКИ
Application.DoCmd (1312)

'---очищаем мусор
Set ShapeTo = Nothing
Set ShapeFrom = Nothing

End Sub

Private Sub CloneSectionLine(ShapeFromID As Long, ShapeToID As Long)
'Процедура клонирования данных из секции "Line"

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Создаем необхоимое количество строк
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowLine, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowLine, j).Formula
    Next j

End Sub

Private Sub CloneSecEvent(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Event"
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

    On erroro GoTo EX

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowEvent)
        ShapeTo.CellsSRC(visSectionObject, visRowEvent, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowEvent, j).Formula
    Next j
Exit Sub
EX:
    SaveLog Err, "CloneSecEvent"
End Sub

Private Sub CloneSecFill(ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования значений строк для секции "Fill"
'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer   ', CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'---Сверяем наборы строк в обеих фигурах, и в случае отсутствия в них искомых - создаем их.
    For j = 0 To ShapeFrom.RowsCellCount(visSectionObject, visRowMisc)
        ShapeTo.CellsSRC(visSectionObject, visRowMisc, j).Formula = ShapeFrom.CellsSRC(visSectionObject, visRowMisc, j).Formula
    Next j

End Sub




Sub CloneSectionUniverseNames(SectionIndex As Integer, ShapeFromID As Long, ShapeToID As Long)
'Процедура копирования необходимого набора свойств указанной секции из Фигуры(ShapeFrom) в Фигуру(ShapeTo)

'---Объявляем переменные
Dim ShapeFrom As Visio.Shape, ShapeTo As Visio.Shape
Dim RowCountFrom As Integer, RowCountTo As Integer
Dim RowNum As Integer, CellNum As Integer

'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
Set ShapeTo = Application.ActivePage.Shapes.ItemFromID(ShapeToID)

'MsgBox Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Name

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
Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1) 'Application.ActivePage.Shapes.ItemFromID(ShapeFromID)
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




'--------------------------------Получение номера слоя--------------------------------------------------
Function LayerImport(ShapeFromID As Long, ShapeToID As Long) As String
'Функция возвращает номер слоя в текущем документе соответствующего слою в документе мастера
Dim ShapeFrom As Visio.Shape
Dim LayerNumber As Integer, LayerName As String
Dim Flag As Boolean
Dim vsoLayer As Visio.Layer

    On Error GoTo EX
'---Присваиваем объектным переменным Фигуры(ShapeFrom и ShpeTo) в соответствии с индексами
    Set ShapeFrom = Application.Documents("Водоснабжение.vss").Masters(ShapeFromID).Shapes(1)

'---Получаем название слоя соответственно номеру в исходном документе
    LayerName = ShapeFrom.Layer(1).Name

'---Проверяем есть ли в текущем документе слой с таким именем
    For i = 1 To Application.ActivePage.Layers.Count
        If Application.ActivePage.Layers(i).Name = LayerName Then
            Flag = True
        End If
    Next i

'---В соответствии с полученным названием определяем номер слоя в текущем документе
    If Flag = True Then
        LayerNumber = Application.ActivePage.Layers(LayerName).Index
    Else
    '---Создаем новый слой с именем слоя к которому принадлежит исходная фигура
        Set vsoLayer = Application.ActiveWindow.Page.Layers.Add(LayerName)
        vsoLayer.NameU = LayerName
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
        LayerNumber = Application.ActivePage.Layers(LayerName).Index
    End If
        
LayerImport = Chr(34) & LayerNumber - 1 & Chr(34)
Exit Function
EX:
    SaveLog Err, "LayerImport"
End Function

