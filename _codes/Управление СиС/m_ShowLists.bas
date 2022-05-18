Attribute VB_Name = "m_ShowLists"
Option Explicit

Public ctrlOn As Boolean

'-------------------Модуль для отображения списков (Подразделения, Свтолы и т.д.)---------------------
Public Sub ShowUnits()
'Показываем имеющиеся на схеме подразделения
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsMainTechnics(cellval(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---Сортируем
    Set units = SortCol(units, "Prop.ArrivalTime", False)
    
    'Заполняем таблицу  с перечнем техники
    If units.Count > 0 Then
        ReDim myArray(units.Count, 5)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Подразделение"
        myArray(0, 2) = "Позывной"
        myArray(0, 3) = "Модель"
        myArray(0, 4) = "Время прибытия"
        myArray(0, 5) = "Личный состав"
    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.Unit", visUnitsString, "") & cellval(shp, "Prop.Owner", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = cellval(shp, "Prop.Call", visUnitsString, "") & cellval(shp, "Prop.About", visUnitsString, "")  '"Позывной"
            myArray(i, 3) = cellval(shp, "Prop.Model", visUnitsString, "")  '"Модель"
            myArray(i, 4) = Format(cellval(shp, "Prop.ArrivalTime"), "DD.MM.YYYY hh:nn:ss")  '"Время прибытия"
            myArray(i, 5) = cellval(shp, "Prop.PersonnelHave")  '"Личный состав"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;75 pt;100 pt;100 pt;50 pt", "ArrivedUnits", "Техника"
    
End Sub

Public Sub ShowPersonnel()
'Показываем имеющиеся на схеме фигуры подразумевающие работу с ними личного состава
Dim i As Integer
Dim j As Integer
Dim row As Integer
Dim shp As Visio.Shape
'Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm

Dim unitsList As Collection
Dim unitName As String
Dim callsList As String
Dim positions As Collection
Dim unitPositions As Collection
Dim pr As String
    
    
    '---Получаем список имеющихся пожарных частей
    A.Refresh Application.ActivePage.Index
    Set unitsList = SortCol(A.GFSShapes, "Prop.Unit", False, visUnitsString)
    Set unitsList = GetUniqueVals(unitsList, _
                                 "Prop.Unit", , " ", " ")

    '---Получаем список всех боевых позиций
    Set positions = FilterShapes(A.GFSShapes, "Prop.PersonnelHave;Prop.Personnel")
    '---Сортируем
    Set positions = SortCol(positions, "Prop.ArrivalTime;Prop.LineTime;Prop.SetTime;Prop.FormingTime;Prop.SquareTime;Prop.FireTime", False, visDate)
    
    'Заполняем таблицу  с перечнем техники
        '---Создаем новый массив для дальнейшего формирования списка
        ReDim myArray(unitsList.Count + positions.Count, 4)
        '---Вставка первой записи
        row = 0
        myArray(row, 0) = "ID"
        myArray(row, 1) = "Тип"
        myArray(row, 2) = "Пожарных"
        myArray(row, 3) = "Работает"
        myArray(row, 4) = "Время"
        
        '---Перебираем названия подразделений и для каждого из них заполняем перечень имеющихся боевых позиций
        For i = 1 To unitsList.Count
            row = row + 1
            unitName = unitsList(i)
            
            '---Получаем список позывных техники данного подразделения
            callsList = StrColToStr(GetUniqueVals( _
                                        FilterShapesAnd(A.GFSShapes, "Prop.PersonnelHave:;Prop.Unit:" & unitName), _
                                        "Prop.Call", , , " "), ", ")
            '---Получаем перечень боевых позиций для данного подразделения
            Set unitPositions = FilterShapes(positions, "Prop.Unit:" & unitName)
            
            
            myArray(row, 0) = -1    '(-1 признак того, что для данной записи нет фигур и переходить к ним не нужно)
            myArray(row, 1) = unitName & ":   " & callsList                      '"Подразделение"
            myArray(row, 2) = GetPersonnelCount(unitPositions)                   '"Боевой расчет"
            myArray(row, 3) = CellSum(unitPositions, "Prop.Personnel")           '"Работает л/с"
            myArray(row, 4) = " "                                                '"Время"
            
            For Each shp In unitPositions
                row = row + 1

                myArray(row, 0) = shp.ID
                myArray(row, 1) = "  " & ChrW(9500) & " " & cellval(shp, "User.IndexPers.Prompt", visUnitsString)  '"Тип"
                myArray(row, 2) = GetPersonnelCount(shp)                          '"Боевой расчет"
                myArray(row, 3) = cellval(shp, "Prop.Personnel", , " ")           '"Работает л/с"
                myArray(row, 4) = Format(pf_GetTime(shp), "DD.MM.YYYY hh:nn")     '"Время"

            Next shp
            
            'Заменяем префикс для последней дочерней записи, при условии таковых
            If unitPositions.Count > 0 Then
                myArray(row, 1) = "  " & ChrW(9492) & Right(myArray(row, 1), Len(myArray(row, 1)) - 3)
            End If
        Next i
    

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;125 pt;75 pt;75 pt;100 pt", "Personnel", "Личный состав"
    
End Sub

Private Function GetPersonnelCount(ByRef shps As Variant) As String
'Считаем количество пожарных (без учета водителей)
Dim shp As Visio.Shape
Dim tmpVal As Integer
Dim tmpSum As Integer
    
    If TypeName(shps) = "Shape" Then        'Если фигура
        Set shp = shps
        tmpVal = cellval(shp, "Prop.PersonnelHave")
        If tmpVal = 0 Then
            tmpSum = tmpVal
        Else
            tmpSum = tmpVal - 1
        End If
    ElseIf TypeName(shps) = "Shapes" Or TypeName(shps) = "Collection" Then     'Если коллекция
        For Each shp In shps
            tmpVal = cellval(shp, "Prop.PersonnelHave")
            If tmpVal > 0 Then
                tmpSum = tmpSum + tmpVal - 1
            End If
        Next shp
    End If
    
    If tmpSum = 0 Then
        GetPersonnelCount = " "
    Else
        GetPersonnelCount = CStr(tmpSum)
    End If
End Function


Public Sub ShowNozzles()
'Показываем имеющиеся на схеме стволы
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsStvols(cellval(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---Сортируем
    Set units = SortCol(units, "Prop.SetTime", False)
    
    'Заполняем таблицу  с перечнем техники
    If units.Count > 0 Then
        ReDim myArray(units.Count, 7)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Подразделение"
        myArray(0, 2) = "Тип ствола"
        myArray(0, 3) = "Позывной"
        myArray(0, 4) = "Время подачи"
        myArray(0, 5) = "Личный состав"
        myArray(0, 6) = "Работа"
        myArray(0, 7) = "Производительность"
    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = cellval(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"Тип ствола"
            myArray(i, 3) = cellval(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 4) = Format(cellval(shp, "Prop.SetTime"), "DD.MM.YYYY hh:nn:ss")  '"Время подачи"
            myArray(i, 5) = cellval(shp, "Prop.Personnel")  '"Личный состав"
            myArray(i, 6) = cellval(shp, "Prop.UseDirection", visUnitsString, "")  '"Работа"
            myArray(i, 7) = cellval(shp, "User.PodOut")  '"Производительность"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt;50 pt;50 pt;50 pt", "Nozzles", "Стволы"
    
End Sub

Public Sub ShowGDZS()
'Показываем имеющиеся на схеме элементы ГДЗС
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsGDZS(cellval(shp, "User.IndexPers")) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    '---Сортируем
    Set units = SortCol(units, "Prop.FormingTime", False)
    
    'Заполняем таблицу  с перечнем техники
    If units.Count > 0 Then
        ReDim myArray(units.Count, 6)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Подразделение"
        myArray(0, 2) = "Тип"
        myArray(0, 3) = "Позывной"
        myArray(0, 4) = "Время формирования"
        myArray(0, 5) = "Личный состав"
        myArray(0, 6) = "СИЗОД"
    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = cellval(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"Тип"
            myArray(i, 3) = cellval(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 4) = Format(cellval(shp, "Prop.FormingTime"), "DD.MM.YYYY hh:nn:ss")  '"Время формирования"
            myArray(i, 5) = cellval(shp, "Prop.Personnel")  '"Личный состав"
            myArray(i, 6) = cellval(shp, "Prop.AirDevice", visUnitsString, " ")  '"СИЗОД"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt;50 pt;50 pt", "GDZS", "ГДЗС"
    
End Sub

Public Sub ShowTimeLine()
'Показываем имеющиеся на схеме элементы ГраФиС
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If pf_IsTimeLine(shp) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
'    Set units = A.Refresh(Application.ActivePage.Index).GFSShapes
    
    '---Сортируем
    Set units = SortCol(units, "Prop.ArrivalTime;Prop.LineTime;Prop.SetTime;Prop.FormingTime;Prop.SquareTime;Prop.FireTime", False, visDate)
    
    'Заполняем таблицу  с перечнем техники
    If units.Count > 0 Then
        ReDim myArray(units.Count, 4)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Подразделение"
        myArray(0, 2) = "Позывной"
        myArray(0, 3) = "Тип"
        myArray(0, 4) = "Время"

    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = cellval(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 3) = cellval(shp, "User.IndexPers.Prompt", visUnitsString)  '"Тип"
            myArray(i, 4) = Format(pf_GetTime(shp), "DD.MM.YYYY hh:nn:ss")  '"Время"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;120 pt;100 pt;100 pt", "TimeLine", "Таймлайн"
    
End Sub

Public Sub ShowStatists()
'Показываем список имеющихся на схеме статистов
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexPers.ipStatist) Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub

    
    'Заполняем таблицу  с перечнем техники
    If units.Count > 0 Then
        ReDim myArray(units.Count, 3)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Состояние"
        myArray(0, 2) = "Информация"
        myArray(0, 3) = "Количество людей"

    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.State", visUnitsString, "")   '"Состояние"
            myArray(i, 2) = cellval(shp, "Prop.Info", visUnitsString, "")  '"Информация"
            myArray(i, 3) = cellval(shp, "Prop.StatistsQuatity")  '"Количество людей"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;100 pt;500 pt;50 pt", "Statists", "Статисты"
    
End Sub

Public Sub ShowExplication()
'Показываем список имеющихся на схеме помещений (Экспликацию)
Dim i As Integer
Dim shp As Visio.Shape
Dim units As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по времени
    Set units = New Collection
    For Each shp In Application.ActivePage.Shapes
        If cellval(shp, "User.ShapeType") = 38 Then
            AddUniqueCollectionItem units, shp
        End If
    Next shp
    If units.Count = 0 Then Exit Sub
    
    'Заполняем таблицу  с перечнем сведений о местах
    If units.Count > 0 Then
        ReDim myArray(units.Count, 5)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "Код"
        myArray(0, 2) = "Назначение"
        myArray(0, 3) = "Имя"
        myArray(0, 4) = "Площадь"
        myArray(0, 5) = "Рассчетное число людей"

    
        For i = 1 To units.Count
            '---Вставка остальных записей
            Set shp = units(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.LocationID", , "")   '"Код"
            myArray(i, 2) = cellval(shp, "Prop.Use", visUnitsString, "")  '"Назначение"
            myArray(i, 3) = cellval(shp, "Prop.Name", visUnitsString, "")  '"Имя"
            myArray(i, 4) = cellval(shp, "Prop.visArea", visUnitsString, "")    '"Площадь"
            myArray(i, 5) = cellval(shp, "Prop.OccupantCount")  '"Рассчетное число людей"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;200 pt;200 pt;100 pt;100 pt", "Places", "Экспликация"
    
End Sub

Public Sub ShowEvacNodes()
'Показываем список имеющихся на схеме узлов эвакуации по результатам расчета
Dim i As Integer
Dim shp As Visio.Shape
Dim nodes As Collection

Dim myArray As Variant
Dim f As frm_ListForm
    
    
    
    '---Формируем коллекцию фигур и сортируем их по номеру
    Set nodes = New Collection
    For Each shp In A.Refresh(Application.ActivePage.Index).GFSShapes
        If IsGFSShapeWithIP(shp, indexPers.ipEvacNode) Then
            AddOrderedNodeItem nodes, shp
        End If
    Next shp
    If nodes.Count = 0 Then Exit Sub
    
    
    'Заполняем таблицу  с перечнем техники
    If nodes.Count > 0 Then
        ReDim myArray(nodes.Count, 9)
        '---Вставка первой записи
        myArray(0, 0) = "ID"
        myArray(0, 1) = "№"
        myArray(0, 2) = "Класс"
        myArray(0, 3) = "Тип"
        myArray(0, 4) = "Ширина"
        myArray(0, 5) = "Длина"
        myArray(0, 6) = "Людей"
        myArray(0, 7) = "Людской поток"
        myArray(0, 8) = "Время участка"
        myArray(0, 9) = "Время общее"
        

        For i = 1 To nodes.Count
            '---Вставка остальных записей
            Set shp = nodes(i)
            myArray(i, 0) = shp.ID
            myArray(i, 1) = cellval(shp, "Prop.NodeNumber", visUnitsString, "")
            myArray(i, 2) = cellval(shp, "Prop.WayClass", visUnitsString, "")
            myArray(i, 3) = cellval(shp, "Prop.WayType", visUnitsString, "")
            myArray(i, 4) = cellval(shp, "Prop.WayWidth", visUnitsString, "")
            myArray(i, 5) = cellval(shp, "Prop.WayLen", visUnitsString, "")
            myArray(i, 6) = cellval(shp, "Prop.PeopleHere", visUnitsString, "")
            myArray(i, 7) = cellval(shp, "Prop.PeopleFlow", visUnitsString, "")
            myArray(i, 8) = cellval(shp, "Prop.tHere", visUnitsString, "")
            myArray(i, 9) = cellval(shp, "Prop.t_Flow", visUnitsString, "")
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;25 pt;100 pt;100 pt;50 pt;50 pt;50 pt;70 pt;70 pt;70 pt", "EvacNodes", "Узлы пути"

End Sub

Private Sub AddOrderedNodeItem(ByRef nodes As Collection, ByVal nodeItem As Visio.Shape)
Dim nextNode As Visio.Shape
    
    Set nextNode = FindHigherNode(nodes, nodeItem)
    'Основная коллекция
    If nextNode Is Nothing Then
        nodes.Add nodeItem, CStr(nodeItem.ID)
    Else
        nodes.Add nodeItem, CStr(nodeItem.ID), CStr(nextNode.ID)
    End If
    
End Sub
Private Function FindHigherNode(ByRef nodes As Collection, ByRef nodeIn As Visio.Shape) As Visio.Shape
'Возвращает элемент в коллекции nodes с номером больше чем у текущего (тот элемент перед которым нужно будет вставить новый)
Dim node As Visio.Shape
Dim nodeNumber As Integer
Dim nodeInNumber As Integer
    
    nodeInNumber = cellval(nodeIn, "Prop.NodeNumber")
    For Each node In nodes
        nodeNumber = cellval(node, "Prop.NodeNumber")
        If nodeNumber > nodeInNumber Then
            Set FindHigherNode = node
            Exit Function
        End If
    Next node
End Function

'----------------------------Функции проверки типов фигур
Private Function pf_IsMainTechnics(ByVal a_IndexPers As Integer) As Boolean
'Является ли индекс индексом техники
    If a_IndexPers <= 20 Or a_IndexPers = 24 Or a_IndexPers = 25 Or a_IndexPers = 26 Or a_IndexPers = 27 Or _
        a_IndexPers = 28 Or a_IndexPers = 29 Or a_IndexPers = 30 Or a_IndexPers = 31 Or a_IndexPers = 32 Or _
        a_IndexPers = 33 Or a_IndexPers = 73 Or a_IndexPers = 74 Or _
        a_IndexPers = 160 Or a_IndexPers = 161 Or a_IndexPers = 162 Or a_IndexPers = 163 Or _
        a_IndexPers = 3000 Or a_IndexPers = 3001 Or a_IndexPers = 3002 Then
        pf_IsMainTechnics = True
    Else
        pf_IsMainTechnics = False
    End If
End Function

Private Function pf_IsStvols(ByVal a_IndexPers As Integer) As Boolean
'Является ли индекс индексом стволов
    If a_IndexPers >= 34 And a_IndexPers <= 39 Or a_IndexPers = 45 Or a_IndexPers = 76 Or a_IndexPers = 77 Then
        pf_IsStvols = True
    Else
        pf_IsStvols = False
    End If
End Function

Private Function pf_IsGDZS(ByVal a_IndexPers As Integer) As Boolean
'Является ли индекс индексом ГДЗС
    If a_IndexPers >= 46 And a_IndexPers <= 48 Or a_IndexPers = 90 Then
        pf_IsGDZS = True
    Else
        pf_IsGDZS = False
    End If
End Function

Private Function pf_IsTimeLine(ByRef a_Shape As Visio.Shape) As Boolean
'Является ли индекс индексом фигур таймлайна
    If a_Shape.CellExists("Prop.ArrivalTime", 0) = True Or a_Shape.CellExists("Prop.LineTime", 0) = True _
        Or a_Shape.CellExists("Prop.SetTime", 0) = True Or a_Shape.CellExists("Prop.FormingTime", 0) = True _
        Or a_Shape.CellExists("Prop.SquareTime", 0) = True Or a_Shape.CellExists("Prop.FireTime", 0) = True _
        Then
        pf_IsTimeLine = True
    Else
        pf_IsTimeLine = False
    End If
End Function

Public Function pf_GetTime(ByRef aO_Shape As Visio.Shape, Optional ByVal default As String = "не определено") As String
'Получение времени фигуры
On Error GoTo ex

    If aO_Shape.CellExists("Prop.ArrivalTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.ArrivalTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.LineTime", 0) = True Then
        pf_GetTime = aO_Shape.Cells("Prop.LineTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SetTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SetTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FormingTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FormingTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.SquareTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.SquareTime").ResultStr(visDate)
        Exit Function
    End If
    If aO_Shape.CellExists("Prop.FireTime", 0) = True Then
            pf_GetTime = aO_Shape.Cells("Prop.FireTime").ResultStr(visDate)
        Exit Function
    End If

Exit Function
ex:
    pf_GetTime = default
End Function
