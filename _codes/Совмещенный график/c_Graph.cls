VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------Класс построения совмещенного графика тушения пожара------------------------------------------
Option Explicit

'---------------------Коллекции-------------------------------------------------------------------------------------
Public ColP_Fires As New Collection
Public ColP_WStvols As New Collection


'---------------------Свойства--------------------------------------------------------------------------------------
Public pi_TargetPageIndex As Integer          'Целевой лист подсчета (0-все листы, 1-... - лист с указанным номером)
Private pcpO_BeginFire As Visio.Shape         'Фигура очага пожара - отчетная точка по времени
Private pcpD_BeginDate As Date                'Время начала пожара
Private FiresColSorted As Boolean             'Сортирована ли уже коллекция плоащдей горения
Private WStvolsColSorted As Boolean           'Сортирована ли уже коллекция водяных стволов

'---Постоянные индексов графиков
Const csi_FireIndex = 3
Const csi_ExtIndex = 4
Const csi_PodIndex = 5
Const csi_PodEffIndex = 6

'---Постоянна версии наборов
Const CP_GrafisVersion = 1

'Класс обрабатывает данные на указанном листе и возвращает массивы данных с результатами анализа
'На вход:
'-номер листа
'-тип графика

'На выход
'-Массивы данных для построения графиков
'-время начала пожара
'-Интенсивность
'-Время окончания тушения




'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
'MsgBox "Класс инициализирован"

End Sub


'----------------------------------------Внешние команды (процедуры и функции к ним)-------------------------------
Public Sub sC_ColRefresh() '(as_PageIndex As Integer)
'Процедура обновления коллекций класса
'Перебирает все элементы находящиеся на указанном as_PageIndex листе
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer
Dim vsO_TargetPage As Visio.Page
Dim tempTime1 As Date
Dim tempTime2 As Date

    On Error GoTo EX

'---Указываем целевую страницу для анализа
    Set vsO_TargetPage = Application.ActiveDocument.Pages(pi_TargetPageIndex)

'---Перебираем все фигуры и в случае если фигура является фигурой ГраФиС анализируем её
For Each vsO_Shape In vsO_TargetPage.Shapes
    If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then 'Является ли фигура фигурой ГраФиС
        If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then 'Проверяем версию фигуры
            vsi_ShapeIndex = vsO_Shape.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
            '---Определяем к какому именно типу относится фигура и выполняем необходимое действие
            Select Case vsi_ShapeIndex
            Case Is = 64 'Площадь пожара
                ColP_Fires.Add vsO_Shape, str(vsO_Shape.ID)
            Case Is = 34 'Ручной водяной
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 36 'Лафетный водяной
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 39 'Лафетный водяной возимый
                If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                    ColP_WStvols.Add vsO_Shape, str(vsO_Shape.ID)
                End If
            Case Is = 70 'Очаг пожара
                Set pcpO_BeginFire = vsO_Shape
                pcpD_BeginDate = pcpO_BeginFire.Cells("Prop.FireTime")
            End Select
        End If
    End If
Next vsO_Shape

'Указываем, что коллекции нуждаются в сортировке
FiresColSorted = False
WStvolsColSorted = False

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "sC_ColRefresh"
End Sub

'------------------------------Процедуры возвращения данных-------------------------------------------------
Public Sub PS_GetFireSquares(ByRef MainArray())
'Процедура получает динамический массив и заполняет его данными о площадах пожара
'Первая цифра (0-абс.время; 1-значение площади), Вторая значение соответсвующего атрибута точки
Dim CurShape As Visio.Shape
Dim curTime As Date
Dim curSquare As Double
Dim i As Integer

    On Error GoTo EX

'---Сортируем фигуры в коллекции площадей горения
    sC_SortFires

'---Переопределяем массив в соответсвии с количеством фигур в коллекции
    If ColP_Fires.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_Fires.Count)
    End If

'---Перебираем все фигуры в коллекции
'    i = 0
'    For Each CurShape In ColP_Fires
'        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
'        MainArray(1, i) = CurShape.Cells("User.FireSquare")
'        i = i + 1
'    Next CurShape
    i = 0
    For Each CurShape In ColP_Fires
        If curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate)) Then
            curSquare = curSquare + CurShape.Cells("User.FireSquare")
        Else
            curSquare = CurShape.Cells("User.FireSquare")
            curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
            i = i + 1
        End If
        MainArray(0, i) = curTime
        MainArray(1, i) = curSquare
        
    Next CurShape
    
'---Удаляем пустые строки из конца массива (если таковые имеются)
    For i = UBound(MainArray, 2) To 1 Step -1
        If Not IsEmpty(MainArray(1, i)) Then Exit For
    Next i
    ReDim Preserve MainArray(1, i)

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_GetFireSquares"
End Sub

Public Sub PS_GetExtSquares(ByRef MainArray())
'Процедура получает динамический массив и заполняет его данными о площадах ТУШЕНИЯ пожара
'Первая цифра (0-абс.время; 1-значение площади), Вторая значение соответсвующего атрибута точки
Dim CurShape As Visio.Shape
Dim curTime As Date
Dim curSquare As Double
Dim i As Integer

    On Error GoTo EX

'---Сортируем фигуры в коллекции площадей горения
    sC_SortFires

'---Переопределяем массив в соответсвии с количеством фигур в коллекции
    If ColP_Fires.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_Fires.Count)
    End If

'---Перебираем все фигуры в коллекции
'    i = 0
'    For Each CurShape In ColP_Fires
'        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
'        MainArray(1, i) = CurShape.Cells("User.ExtSquare")
'        i = i + 1
'    Next CurShape
    i = 0
    For Each CurShape In ColP_Fires
        If curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate)) Then
            curSquare = curSquare + CurShape.Cells("User.ExtSquare")
        Else
            curSquare = CurShape.Cells("User.ExtSquare")
            curTime = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SquareTime").ResultStr(visDate))
            i = i + 1
        End If
        MainArray(0, i) = curTime
        MainArray(1, i) = curSquare
        
    Next CurShape
    
'---Удаляем пустые строки из конца массива (если таковые имеются)
    For i = UBound(MainArray, 2) To 1 Step -1
        If Not IsEmpty(MainArray(1, i)) Then Exit For
    Next i
    ReDim Preserve MainArray(1, i)

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_GetExtSquares"
End Sub

Public Sub PS_GetWStvolsPodOut(ByRef MainArray())
'Процедура получает динамический массив и заполняет его данными о расходах водяных стволов
'Первая цифра (0-абс.время; 1-значение расхода), Вторая значение соответсвующего атрибута точки
Dim CurShape As Visio.Shape
Dim i As Integer

    On Error GoTo EX

'---Сортируем фигуры в коллекции водяных стволов
    sC_SortWStvols

'---Переопределяем массив в соответсвии с количеством фигур в коллекции
    If ColP_WStvols.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_WStvols.Count - 1)
    End If

'---Перебираем все фигуры в коллекции
    i = 0
    For Each CurShape In ColP_WStvols
        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SetTime").ResultStr(visDate))
        MainArray(1, i) = CurShape.Cells("User.PodOut")
        i = i + 1
    Next CurShape

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_GetWStvolsPodOut"
End Sub

Public Sub PS_GetWStvolsEffPodOut(ByRef MainArray())
'Процедура получает динамический массив и заполняет его данными о ЭФФЕКТИВНЫХ расходах водяных стволов
'Первая цифра (0-абс.время; 1-значение расхода), Вторая значение соответсвующего атрибута точки
Dim CurShape As Visio.Shape
Dim i As Integer

    On Error GoTo EX

'---Сортируем фигуры в коллекции водяных стволов
    sC_SortWStvols

'---Переопределяем массив в соответсвии с количеством фигур в коллекции
    If ColP_WStvols.Count = 0 Then
        Exit Sub
    Else
        ReDim MainArray(1, ColP_WStvols.Count - 1)
    End If
    
'---Перебираем все фигуры в коллекции
    i = 0
    For Each CurShape In ColP_WStvols
        MainArray(0, i) = DateDiff("s", pcpD_BeginDate, CurShape.Cells("Prop.SetTime").ResultStr(visDate))
        MainArray(1, i) = CurShape.Cells("User.PodOutEff")
        i = i + 1
    Next CurShape

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_GetWStvolsEffPodOut"
End Sub

Public Function PF_GetBeginDateTime() As Date
'Функция возвращает значения географического времени начала пожара
    PF_GetBeginDateTime = pcpD_BeginDate
End Function

Public Function PF_GetIntence(ByVal ShpPos As Integer) As Single
'Функция возвращает значение интенсивности фигуры площади пожара в списке согласно позиции ShpPos
On Error GoTo EX
    
'    If ColP_Fires.Count = 0 Then Exit Function
    
'---Сортируем фигуры в коллекции площадей пожара
    sC_SortFires
'---Получаем данные
    PF_GetIntence = ColP_Fires(ShpPos).Cells("User.WaterIntense")

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetIntence"
    PF_GetIntence = 0.1
End Function

Public Function PF_GetTimeEnd(ByVal Modificator As Single, ByVal Opt As String) As Long
'Функция возвращает значение времени прошедшего с момента начал пожара до момент окончания
'Время вычисляетс путем умножения времени максимальной площади пожара на модификатор
'Единицы выбираются в указанных опцией
Dim i As Integer
    
'---Сортируем фигуры в коллекции площадей пожара
    sC_SortFires
    
    PF_GetTimeEnd = DateDiff(Opt, pcpD_BeginDate, ColP_Fires(ColP_Fires.Count).Cells("Prop.SquareTime").ResultStr(visDate)) * Modificator
End Function

Public Function PF_CheckFireBeginExist() As Boolean
'Функция возвращает Ложь, если фигуры очага нет на листе
    PF_CheckFireBeginExist = fC_FireCheck(pi_TargetPageIndex)
End Function

Public Function GetMaxGraphSize(ByVal ShpPos) As Single
Dim maxSquare As Single
Dim exp As Single
    
    maxSquare = PF_GetMaxSquare(ShpPos)
    exp = PF_GetExpence / PF_GetIntence(ShpPos)
        
    If exp > maxSquare Then
        GetMaxGraphSize = exp
    Else
        GetMaxGraphSize = maxSquare
    End If
End Function

Public Function PF_GetMaxSquare(ByVal ShpPos) As Single
'Функция возвращает значение максимальной площади пожара
Dim i As Integer
    
On Error GoTo EX
    
    If ColP_Fires.Count = 0 Then Exit Function
    
'---Сортируем фигуры в коллекции площадей пожара
    sC_SortFires
'---Получаем данные
    PF_GetMaxSquare = RUp(ColP_Fires(ShpPos).Cells("User.FireSquare"), 1.1)

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetMaxSquare"
    PF_GetMaxSquare = 320
End Function

Public Function PF_GetExpence() As Single
'Функция возвращает значение максимального расхода воды
Dim i As Integer
Dim shp As Visio.Shape
Dim val As Double
    
On Error GoTo EX

'---Получаем данные
    For Each shp In ColP_WStvols
        val = val + shp.Cells("User.PodOut").Result(visNumber)
    Next shp
    PF_GetExpence = val

Exit Function
EX:
    MsgBox Err.description
    SaveLog Err, "PF_GetExpence"
    PF_GetExpence = 32
End Function


'-----------------------------------Процедуры сортировок коллекций по времени--------------------------------
Private Sub sC_SortFires()
'Процедура сортировки элементов в коллекции площадей пожара по возрастанию Времени отсечки
Dim vsi_MinTime As Long
Dim vsi_CurTime As Long
Dim vsCol_TempCol As Collection
Dim i As Integer
Dim k As Integer
Dim j As Integer

    On Error GoTo EX

'Если коллекция уже сортирована - выходим
    If FiresColSorted = True Then Exit Sub

'---Объявляем новую коллекцию
Set vsCol_TempCol = New Collection

'---Запускаем цикл повторений из числа повторений равного изначальному количеству эл-тов в коллекции
For i = 1 To ColP_Fires.Count
    vsi_MinTime = DateDiff("s", pcpD_BeginDate, ColP_Fires.Item(1).Cells("Prop.SquareTime").ResultStr(visDate))
    j = 1
    '---Запускаем цикл повторений из числа повторений равного текущему количеству эл-тов в коллекции
    For k = 1 To ColP_Fires.Count
        vsi_CurTime = DateDiff("s", pcpD_BeginDate, ColP_Fires.Item(k).Cells("Prop.SquareTime").Result(visDate))
        If vsi_CurTime < vsi_MinTime Then 'Если время текущего эл-та меньше установленного минимального, то запоминаем его номер
            vsi_MinTime = vsi_CurTime
            j = k
        End If
    Next k
    '---Добавляем во временный массив эл-т с наименьшим значением времени
    vsCol_TempCol.Add ColP_Fires.Item(j), str(ColP_Fires.Item(j).ID)
    ColP_Fires.Remove j 'Из исходной коллекции - удаляем его
    
Next i

'---Обновляем изначальную коллекцию в соответствии с полученной временной
    Set ColP_Fires = vsCol_TempCol
    Set vsCol_TempCol = Nothing
    
'---Указываем, что коллекция уже отсоритрована
    FiresColSorted = True

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "sC_SortFires"
Exit Sub
End Sub


Private Sub sC_SortWStvols()
'Процедура сортировки элементов в коллекции площадей пожара по возрастанию Времени отсечки
Dim vsi_MinTime As Integer
Dim vsi_CurTime As Integer
Dim vsCol_TempCol As Collection
Dim i As Integer
Dim k As Integer
Dim j As Integer

    On Error GoTo EX

'Если коллекция уже сортирована - выходим
    If WStvolsColSorted = True Then Exit Sub

'---Объявляем новую коллекцию
Set vsCol_TempCol = New Collection

'---Запускаем цикл повторений из числа повторений равного изначальному количеству эл-тов в коллекции
For i = 1 To ColP_WStvols.Count
    vsi_MinTime = DateDiff("s", pcpD_BeginDate, ColP_WStvols.Item(1).Cells("Prop.SetTime").ResultStr(visDate))
    j = 1
    '---Запускаем цикл повторений из числа повторений равного текущему количеству эл-тов в коллекции
    For k = 1 To ColP_WStvols.Count
        vsi_CurTime = DateDiff("s", pcpD_BeginDate, ColP_WStvols.Item(k).Cells("Prop.SetTime").ResultStr(visDate))
        If vsi_CurTime < vsi_MinTime Then 'Если время текущего эл-та меньше установленного минимального, то запоминаем его номер
            vsi_MinTime = vsi_CurTime
            j = k
        End If
    Next k
    '---Добавляем во временный массив эл-т с наименьшим значением времени
    vsCol_TempCol.Add ColP_WStvols.Item(j), str(ColP_WStvols.Item(j).ID)
    ColP_WStvols.Remove j 'Из исходной коллекции - удаляем его
    
Next i

'---Обновляем изначальную коллекцию в соответствии с полученной временной
    Set ColP_WStvols = vsCol_TempCol
    Set vsCol_TempCol = Nothing
    
'---Указываем, что коллекция уже отсоритрована
    WStvolsColSorted = True

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "sC_SortWStvols"
End Sub


'-------------------------------------------------Служебные функции-----------------------------------------------------
Private Function fC_MaxValue(afi_ColIndex As Byte, afs_CellName As String) As Variant
'Функция возвращает максимальное значение указанной ячейки в указанной коллекции
Dim vfO_TempCol As Collection
Dim vfO_Shape As Visio.Shape
Dim vfs_MaxValue As Single

    Set vfO_TempCol = New Collection
    vfs_MaxValue = 0

    Select Case afi_ColIndex
        Case Is = 1
            Set vfO_TempCol = ColP_Fires
        Case Is = 2
            Set vfO_TempCol = ColP_WStvols
    End Select

    For Each vfO_Shape In vfO_TempCol
        If vfO_Shape.Cells(afs_CellName).Result(visNumber) > vfs_MaxValue Then _
            vfs_MaxValue = vfO_Shape.Cells(afs_CellName).Result(visNumber)
    Next vfO_Shape

Set vfO_TempCol = Nothing
fC_MaxValue = vfs_MaxValue
End Function


Private Function fC_FireCheck(as_PageIndex As Integer) As Boolean
'Функция проверки наличия фигуры очага пожара на листе
Dim vfO_Page As Visio.Page
Dim vfO_Shape As Visio.Shape

fC_FireCheck = False

    Set vfO_Page = Visio.Application.ActiveDocument.Pages(as_PageIndex)
    
    For Each vfO_Shape In vfO_Page.Shapes
        If vfO_Shape.CellExists("User.IndexPers", 0) = True Then
            Debug.Print vfO_Shape.Name & " - " & vfO_Shape.Cells("User.IndexPers").Result(visNumber)
            If vfO_Shape.Cells("User.IndexPers").Result(visNumber) = 70 Then
                fC_FireCheck = True
                Set vfO_Page = Nothing
                Exit Function
            End If
        End If
    Next vfO_Shape
    
    Set vfO_Page = Nothing
    
fC_FireCheck = False
End Function






