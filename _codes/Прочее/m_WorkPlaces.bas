Attribute VB_Name = "m_WorkPlaces"
Option Explicit




Public Sub PS_AddWorkPlaces(ShpObj As Visio.Shape)
'Процедура добавляет фигуры рабочих мест в соответствии с комнатами
Dim WorkPlaceBuilder As c_WorkPlaces

'---Проверяем наличие трафарета WALL_M.VSS
     If PF_DocumentOpened("WALL_M.VSS") = False Then ' And PF_DocumentOpened("WALL_M.VSSX") = False Then
        ShpObj.Delete
        MsgBox "Трафарет 'Структурные элементы' не подключен! Выполнение функции невозможно!'", vbCritical
        Exit Sub
     End If
    
'---Вбрасываем фигуры
    Set WorkPlaceBuilder = New c_WorkPlaces
        WorkPlaceBuilder.S_SetFullShape
    Set WorkPlaceBuilder = Nothing
    
'---Удаляем стартовую фигуру
    ShpObj.Delete

End Sub

Public Sub PS_WorkPlacesRenum(ShpObj As Visio.Shape)
'Процедура перебирает все имеющиеся на листе фигуры рабочих мест и перенумеровывает их
Dim WorkPlace As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim i As Integer

'---Выбираем все фигуры слоя "Рабочее место"
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Место")
    Application.ActiveWindow.Selection = vsoSelection
    
'---Перебираем и перенумеровываем фигуры фигуры
    i = 1
    For Each WorkPlace In vsoSelection
        WorkPlace.Cells("Prop.LocationID").FormulaU = i
        i = i + 1
    Next
    
'---Удаляем стартовую фигуру
    ShpObj.Delete

End Sub

'-----------------------------------Вставка экспликаций-----------------------------------------
Public Sub PS_AddExplicationTable(ShpObj As Visio.Shape)
'Прока добавляет таблицу экспликации
Dim vO_Shape As Visio.Shape
Dim vO_NewString As Visio.Shape
Dim vO_StringMaster As Visio.Master
Dim i As Integer

Dim colWorkplaces As Collection

On Error GoTo Tail
    
    '---Создаем коллекцию мест
    Set colWorkplaces = New Collection
    '---Наполняем коллекцию мест
    For Each vO_Shape In Application.ActivePage.Shapes
        If PFB_isPlace(vO_Shape) Then
            colWorkplaces.Add vO_Shape
        End If
    Next vO_Shape
    
    '---Если в коллекции нет фигур - Выходим
        If colWorkplaces.Count = 0 Then Exit Sub
    
    '---Сортируем коллекцию
    sC_SortPlaces colWorkplaces
    
    '---Отключаем учет событий приложением Visio
    Application.EventsEnabled = False
    
    '---Получаем мастер строк
    Set vO_StringMaster = ThisDocument.Masters("Экспликация")
    
    '---Перебираем все фигуры на листе, и если фигура - фигура места, создаем строки экспликации
    i = 0
    For Each vO_Shape In colWorkplaces
        If PFB_isPlace(vO_Shape) Then
            i = i + 1
            
            If i = 1 Then
                '---Присваиваем значения имеющейся фигуре
                ShpObj.Cells("User.PlaceSheetName").FormulaU = """" & vO_Shape.NameID & """"
            Else
                '---Добавляем новую строчку
                Set vO_NewString = Application.ActivePage.Drop(vO_StringMaster, _
                                    ShpObj.Cells("PinX").Result(visInches), _
                                    ShpObj.Cells("PinY").Result(visInches) - ShpObj.Cells("Height").Result(visInches) * (i - 1))
                vO_NewString.Cells("User.PlaceSheetName").FormulaU = """" & vO_Shape.NameID & """"
            End If
        End If
    Next vO_Shape

    '---Заключительный этап
    Application.EventsEnabled = True
    Set colWorkplaces = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_AddExplicationTable"
    Application.EventsEnabled = True
    Set colWorkplaces = Nothing
End Sub


Private Sub sC_SortPlaces(ByRef aC_PlacesCollection As Collection)
'Процедура сортировки элементов в коллекции площадей пожара по возрастанию Времени отсечки
Dim vsi_MinKod, vsi_CurKod As Integer
'Dim vO_Place1 As Visio.Shape
'Dim vO_Place2 As Visio.Shape
Dim vsCol_TempCol As Collection
Dim i, k, j As Integer

'---Объявляем новую коллекцию
Set vsCol_TempCol = New Collection

'---Запускаем цикл повторений из числа повторений равного изначальному количеству эл-тов в коллекции
For i = 1 To aC_PlacesCollection.Count
    vsi_MinKod = aC_PlacesCollection.Item(1).Cells("Prop.LocationID").Result(visNumber)
    j = 1
    '---Запускаем цикл повторений из числа повторений равного текущему количеству эл-тов в коллекции
    For k = 1 To aC_PlacesCollection.Count
        vsi_CurKod = aC_PlacesCollection.Item(k).Cells("Prop.LocationID").Result(visNumber)
        If vsi_CurKod < vsi_MinKod Then 'Если текущий код меньше минимального, то запоминаем номер позиции
            vsi_MinKod = vsi_CurKod
            j = k
        End If
    Next k
    '---Добавляем во временный массив эл-т с наименьшим значением времени
    vsCol_TempCol.Add aC_PlacesCollection.Item(j), str(aC_PlacesCollection.Item(j).ID)
    aC_PlacesCollection.Remove j 'Из исходной коллекции - удаляем его
    
Next i


'---Обновляем изначальную коллекцию в соответствии с полученной временной
Set aC_PlacesCollection = vsCol_TempCol

Set vsCol_TempCol = Nothing
End Sub





