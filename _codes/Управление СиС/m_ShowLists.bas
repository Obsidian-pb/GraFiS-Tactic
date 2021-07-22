Attribute VB_Name = "m_ShowLists"
Option Explicit


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
        If pf_IsMainTechnics(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "") & CellVal(shp, "Prop.Owner", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = CellVal(shp, "Prop.Call", visUnitsString, "") & CellVal(shp, "Prop.About", visUnitsString, "")  '"Позывной"
            myArray(i, 3) = CellVal(shp, "Prop.Model", visUnitsString, "")  '"Модель"
            myArray(i, 4) = Format(CellVal(shp, "Prop.ArrivalTime"), "hh:mm:ss")  '"Время прибытия"
            myArray(i, 5) = CellVal(shp, "Prop.PersonnelHave")  '"Личный состав"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;75 pt;75 pt;100 pt;100 pt;50 pt", "ArrivedUnits", "Техника"
    
End Sub

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
        If pf_IsStvols(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"Тип ствола"
            myArray(i, 3) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 4) = Format(CellVal(shp, "Prop.SetTime"), "hh:mm:ss")  '"Время подачи"
            myArray(i, 5) = CellVal(shp, "Prop.Personnel", "")  '"Личный состав"
            myArray(i, 6) = CellVal(shp, "Prop.UseDirection", "")  '"Работа"
            myArray(i, 7) = CellVal(shp, "User.PodOut", "")  '"Производительность"
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
        If pf_IsGDZS(CellVal(shp, "User.IndexPers")) Then
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString, "")  '"Тип"
            myArray(i, 3) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 4) = Format(CellVal(shp, "Prop.FormingTime"), "hh:mm:ss")  '"Время формирования"
            myArray(i, 5) = CellVal(shp, "Prop.Personnel", "")  '"Личный состав"
            myArray(i, 6) = CellVal(shp, "Prop.AirDevice", visUnitsString, " ")  '"СИЗОД"
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
            myArray(i, 1) = CellVal(shp, "Prop.Unit", visUnitsString, "")  '"Подразделение"
            myArray(i, 2) = CellVal(shp, "Prop.Call", visUnitsString, "")  '"Позывной"
            myArray(i, 3) = CellVal(shp, "User.IndexPers.Prompt", visUnitsString)  '"Тип"
            myArray(i, 4) = Format(pf_GetTime(shp), "hh:mm:ss")  '"Время"
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
            myArray(i, 1) = CellVal(shp, "Prop.State", visUnitsString, "")   '"Состояние"
            myArray(i, 2) = CellVal(shp, "Prop.Info", visUnitsString, "")  '"Информация"
            myArray(i, 3) = CellVal(shp, "Prop.StatistsQuatity", , "")  '"Количество людей"
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
        If CellVal(shp, "User.ShapeType") = 38 Then
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
            myArray(i, 1) = CellVal(shp, "Prop.LocationID", , "")   '"Код"
            myArray(i, 2) = CellVal(shp, "Prop.Use", visUnitsString, "")  '"Назначение"
            myArray(i, 3) = CellVal(shp, "Prop.Name", visUnitsString, "")  '"Имя"
            myArray(i, 4) = CellVal(shp, "Prop.visArea", visUnitsString, "")    '"Площадь"
            myArray(i, 5) = CellVal(shp, "Prop.OccupantCount", , "")  '"Рассчетное число людей"
        Next i
    End If

    '---Показываем форму
    Set f = New frm_ListForm
    f.Activate myArray, "0 pt;50 pt;200 pt;200 pt;100 pt;100 pt", "Places", "Экспликация"
    
End Sub




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

Private Function pf_GetTime(ByRef aO_Shape As Visio.Shape) As String
'Получение времени фигуры
On Error GoTo EX

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
EX:
    pf_GetTime = "не определено"
End Function
