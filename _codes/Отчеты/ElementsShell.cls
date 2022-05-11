VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElementsShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================Клас-обертка для коллекций элементов==========================================
Private elemCollection As Collection            'Полная коллекция
Private elemWithCellsCollection As Collection   'Коллекция элементов для которых заданы  имена расчетных ячеек

'-------------------------------------Свойства---------------------------------



'-------------------------------------Базовые процедуры классса------------------------------------
Private Sub Class_Initialize()
    Set elemCollection = New Collection
    Set elemWithCellsCollection = New Collection
    fillElemCollection
End Sub
Private Sub Class_Terminate()
    Set elemCollection = Nothing
    Set elemWithCellsCollection = Nothing
End Sub

'-------------------------------------Свойства класса----------------------------------------------
Public Property Get count() As Integer
'Количество элементов в коллекции
    count = elemCollection.count
End Property




'-------------------------------------Внутренние процедуры и функции классса----------------------------------------------
Private Sub fillElemCollection()
'Заполняем коллекцию элементов
    AddNewElement "RSCHSTechincsHave", "Техники РСЧС имеется", "{$} ({TotalFireCount} - пожарно-спасательной)"
    AddNewElement "MainPAHave", "ПА общего применения имеется", , , , , False
    AddNewElement "TargetedPAHave", "ПА целевого применения имеется", , , , , False
    AddNewElement "MainOverallHave", "ПА основных имеется", "{$} ({TargetedPAHave} - цел.прим., {MainPAHave} - общ.прим.)"
    AddNewElement "PAHighHave", "ПА высотных имеется", , , , , False
    AddNewElement "SpecialPAHave", "ПА специальных имеется", "{$} ({PAHighHave} - высотных)"
    AddNewElement "OtherTechincsHave", "Прочей пожарно-спасательной техники имеется"
    
    AddNewElement "ACCount", "АЦ имеется", , , , , False
    AddNewElement "ACNeed", "АЦ требуется", , , , , False
    AddNewElement "ANRNeed", "АНР требуется", , , , , False
    AddNewElement "AGCount", "АГДЗС имеется", , , , , False
    AddNewElement "ALCount", "АЛ имеется", , , , , False
    AddNewElement "AKPCount", "АКП имеется", , , , , False
    AddNewElement "MVDCount", "Техники МВД имеется", , , , , False
    AddNewElement "SMPCount", "Техники СМП имеется", , , , , False
    AddNewElement "TechnicsNotMchsHave", "Техники прочих служб имеется", "{$} ({MVDCount} - МВД, {SMPCount} - СМП)"
    AddNewElement "TechnicsNotMchsSmpMvdHave", "Техники имеется кроме ПАСР, СМП, МВД"
    AddNewElement "TechnicsNotMchsOtherHave", "Прочей техники имеется"
    AddNewElement "TotalTechCount", "Техники всего имеется", , , , , False
    AddNewElement "TotalFireCount", "Пожарной и аварийно-спасательной техники всего имеется", , , , , False
    
    AddNewElement "FireSquare", "Площадь пожара", "{$} м.кв. ({ExtSquare} м.кв. - площадь тушения)", , "User.FireSquare"
    AddNewElement "ExtSquare", "Площадь тушения", , , "User.ExtSquare", , False
    AddNewElement "FactStreamW", "Расход воды общий, фактический", "{$} л/с", , "User.PodOut"
    AddNewElement "FactStreamWExt", "Расход воды на тушение, фактический", , , , , False
    AddNewElement "FactStreamWProt", "Расход воды на защиту, фактический", , , , , False
    AddNewElement "FactStreamWChill", "Расход воды на охлаждение, фактический", , , , , False
    AddNewElement "IntenseWater", "Интенсивность подачи воды"
    AddNewElement "NeedStreamW", "Расход воды общий, требуемый", "{$} л/с", , "User.NeedStreamW", , True
    
    AddNewElement "PersonnelHave", "Личного состава имеется", "{$} чел. ({PersonnelNeed} чел. - требуется)"
    AddNewElement "PersonnelNeed", "Личного состава требуется", , , "Prop.Personnel", , False
    
    AddNewElement "StvolWHave", "Стволов водяных подано", "{$} ({StvolWBHave} - ств. 'Б', {StvolWAHave} - ств. 'А', {StvolWLHave} - ств. лаф.)"
    AddNewElement "StvolWBHave", "Стволов Б подано", , , , , False
    AddNewElement "StvolWAHave", "Стволов А подано", , , , , False
    AddNewElement "StvolWLHave", "Стволов водяных лафетных подано", , , , , False
    AddNewElement "StvolWBNeed", "Стволов Б требуется", , , , , False
    AddNewElement "StvolWANeed", "Стволов А требуется", , , , , False
    AddNewElement "StvolWLNeed", "Стволов водяных лафетных требуется", , , , , False
    AddNewElement "StvolHave", "Стволов всего подано", "{$} ({StvolExtHave} - тушение, {StvolProtHave} - защита, {StvolChillHave} - охлаждение, {StvolEvacHave} - эвакуация, {StvolCheckHave} - проверка)", , , , True
    AddNewElement "StvolExtHave", "Стволов на тушение, подано", , , , , False
    AddNewElement "StvolProtHave", "Стволов на защиту, подано", , , , , False
    AddNewElement "StvolChillHave", "Стволов на охлаждение, подано", , , , , False
    AddNewElement "StvolEvacHave", "Стволов для эвакуации, подано", , , , , False
    AddNewElement "StvolCheckHave", "Стволов на проверку, подано", , , , , False
    
    AddNewElement "StvolFoamHave", "Стволов пенных подано"
    AddNewElement "StvolGasHave", "Стволов газовых подано"
    AddNewElement "StvolPowderHave", "Стволов порошковых подано"

    AddNewElement "GetedWaterValue", "Забирается воды из внешних источников", "{$} л/с (max = {GetedWaterValueMax} л/с )", , , , False
    AddNewElement "GetedWaterValueMax", "Забирается воды из внешних источников Максимум", , , , , False
    AddNewElement "GetingWaterCount", "МСП установлено на водоисточники", "{$} ({PANeedOnWaterSource} - требуется установить)"
    AddNewElement "PANeedOnWaterSource", "Требуется установить на водоисточники", , , , , False
    AddNewElement "WaterEternal", "Бесконечный запас воды", , False, , False, False

    AddNewElement "GDZSChainsCountWork", "Звеньев ГДЗС работает", "{$} ({GDZSMansCountWork} газодымозащитников)"
    AddNewElement "GDZSChainsCountWorkSoft", "Звеньев ГДЗС работает в стандартных условиях", , , , , False
    AddNewElement "GDZSChainsCountWorkHard", "Звеньев ГДЗС работает в сложных условиях", , , , , False
    AddNewElement "GDZSChainsCountNeed", "Звеньев ГДЗС требуется (+ резерв)"
    AddNewElement "GDZSChainsRezCountHave", "Звеньев ГДЗС в резерве имеется", "{$} ({GDZSMansRezCountHave} газодымозащитников)", , , , False
    AddNewElement "GDZSChainsRezCountNeed", "Звеньев ГДЗС в резерве требуется", , , , , False
'    AddNewElement "GDZSChainsCanBeFormed", "Звеньев ГДЗС Может быть сформировано"
    AddNewElement "GDZSMansHave", "Газодымозащитников имеется на машинах", , , "Prop.GDZSHave", , True
    AddNewElement "GDZSMansCountWork", "Газодымозащитников работает", , , , , False
    AddNewElement "GDZSMansCountNeed", "Газодымозащитников требуется", , , , , False
    AddNewElement "GDZSMansRezCountHave", "Газодымозащитников в резервных звеньях имеется", , , , , False
    AddNewElement "GDZSMansRezCountNeed", "Газодымозащитников в резерве требуется", , , , , False
    AddNewElement "GDZSDiscr", "Несоотвествие количества газодымозащитников в звене сложным условиям", , False
    AddNewElement "GDZSPBCount", "ПБ ГДЗС имеется", "{$} ({GDZSPBNeed} требуется)", , , , True
    AddNewElement "GDZSKPPCount", "КПП ГДЗС имеется", "{$} ({GDZSKPPNeed} требуется)", , , , True
    AddNewElement "GDZSPBNeed", "ПБ ГДЗС требуется", , , , , False
    AddNewElement "GDZSKPPNeed", "КПП ГДЗС требуется", , , , , False

    AddNewElement "Hoses38Count", "Рукавов 38 мм используется", , , , , False
    AddNewElement "Hoses51Count", "Рукавов 51 мм используется", , , , , False
    AddNewElement "Hoses66Count", "Рукавов 66 мм используется", , , , , False
    AddNewElement "Hoses77Count", "Рукавов 77 мм используется", , , , , False
    AddNewElement "Hoses89Count", "Рукавов 89 мм используется", , , , , False
    AddNewElement "Hoses110Count", "Рукавов 110 мм используется", , , , , False
    AddNewElement "Hoses150Count", "Рукавов 150 мм используется", , , , , False
    AddNewElement "Hoses200Count", "Рукавов 200 мм используется", , , , , False
    AddNewElement "Hoses250Count", "Рукавов 250 мм используется", , , , , False
    AddNewElement "Hoses300Count", "Рукавов 300 мм используется", , , , , False
    AddNewElement "HosesCount", "Рукавов используется всего", "{$} (" & _
                                "{Hoses38Count} - 38 мм, {Hoses51Count} - 51 мм, {Hoses66Count} - 66 мм, {Hoses77Count} - 77 мм, {Hoses89Count} - 89 мм, " & _
                                "{Hoses110Count} - 110 мм, {Hoses150Count} - 150 мм, {Hoses200Count} - 200 мм, {Hoses250Count} - 250 мм, {Hoses300Count} - 300 мм" & _
                                ")"
    AddNewElement "HosesLength", "Рукавных линий длина", "{$} м", , , , False

    AddNewElement "Hoses38Have", "Рукавов 38 мм имеется", , , "Prop.Hose38", , False
    AddNewElement "Hoses51Have", "Рукавов 51 мм имеется", , , "Prop.Hose51", , False
    AddNewElement "Hoses66Have", "Рукавов 66 мм имеется", , , "Prop.Hose66", , False
    AddNewElement "Hoses77Have", "Рукавов 77 мм имеется", , , "Prop.Hose77", , False
    AddNewElement "Hoses89Have", "Рукавов 89 мм имеется", , , "Prop.Hose89", , False
    AddNewElement "Hoses110Have", "Рукавов 110 мм имеется", , , "Prop.Hose110", , False
    AddNewElement "Hoses150Have", "Рукавов 150 мм имеется", , , "Prop.Hose150", , False
    AddNewElement "Hoses200Have", "Рукавов 200 мм имеется", , , "Prop.Hose200", , False
    AddNewElement "Hoses250Have", "Рукавов 250 мм имеется", , , "Prop.Hose250", , False
    AddNewElement "Hoses300Have", "Рукавов 300 мм имеется", , , "Prop.Hose300", , False
    AddNewElement "HosesHave", "Рукавов имеется всего", "{$} (" & _
                                "{Hoses38Have} - 38 мм, {Hoses51Have} - 51 мм, {Hoses66Have} - 66 мм, {Hoses77Have} - 77 мм, {Hoses89Have} - 89 мм, " & _
                                "{Hoses110Have} - 110 мм, {Hoses150Have} - 150 мм, {Hoses200Have} - 200 мм, {Hoses250Have} - 250 мм, {Hoses300Have} - 300 мм" & _
                                ")", , , , True
    AddNewElement "HoseSkatka", "Напорных рукавов в скатках", , , , , False
    
    AddNewElement "WaterValueNeed10min", "Воды требуется (на 10мин)", "{$} л"
    AddNewElement "WaterValueHave", "Воды имеется", "{$} л"
    AddNewElement "FoamValueeHave", "Пенообразователя имеется", "{$} л", , "Prop.Foam"

    AddNewElement "BUCount", "БУ создано"
    AddNewElement "SPRCount", "СПР создано"
    AddNewElement "ShtabCount", "Количество штабов", , 0, , False, False
    AddNewElement "RNBDCount", "Количество решающих направлений боевых действий", , 0, , False, False
    AddNewElement "OchagCount", "Количество очагов", , 0, , False, False
    AddNewElement "SpreadCount", "Количество направлдений распространения огня", , 0, , False, False
    AddNewElement "SmokeCount", "Задымление", , 0, , False, False
    AddNewElement "WaterSourceCount", "Водоисточники", , 0, , False, False
    AddNewElement "DistanceCount", "Расстояния", , 0, , False, False
    AddNewElement "DistanceMin", "Расстояние до ближайшего водоисточника", "{$} м"
    AddNewElement "LinesCount", "Отдельные рукавные линии", , 0, , False, False
    AddNewElement "LinesPosCount", "Подписи положения рукавных линий", , 0, , False, False
    AddNewElement "LinesLableCount", "Подписи рукавных линий", , 0, , False, False
    AddNewElement "WorklinesCount", "Количество рабочих линий", , 0, , False, False
    AddNewElement "MaglinesCount", "Количество магистральных линий", , 0, , False, False
    AddNewElement "BuildCount", "Количество зданий", , 0, , False, False
    AddNewElement "SOCount", "Количество степеней огнестойкости", , 0, , False, False
    AddNewElement "FireCount", "Количество площадей горения", , 0, , False, False
    AddNewElement "OrientCount", "Ориентиры", , 0, , False, False
    
    AddNewElement "InjuredCount", "Пострадавших", "{$} чел. ({DeadCount} чел. - погибших)"
    AddNewElement "DeadCount", "Погибших", , , , , False

    AddNewElement "AllHosesWithLabels", "Все рукавные линии с подписями", , True, , False, False
    AddNewElement "AllHosesWithPos", "Все рабочие рукавные линии с подписями положения", , True, , False, False
    
    'Временные показатели
    AddNewElement "FireTime", "Время возникновения пожара", , , "Prop.FireTime", True, False
    AddNewElement "FindTime", "Время обнаружения пожара", , , "Prop.FindTime", True, False
    AddNewElement "InfoTime", "Время сообщения о пожаре", , , "Prop.InfoTime", True, False
    AddNewElement "FirstArrivalTime", "Время прибытия первого подразделения", , , "Prop.FirstArrivalTime", True, False
    AddNewElement "FirstStvolTime", "Время подачи первого ствола", , , "Prop.FirstStvolTime", True, False
    AddNewElement "LocalizationTime", "Время локализации", , , "Prop.LocalizationTime", True, False
    AddNewElement "LOGTime", "Время ликвидации открытого горения", , , "Prop.LOGTime", True, False
    AddNewElement "LPPTime", "Время ликвидации последствий пожара", , , "Prop.LPPTime", True, False
    AddNewElement "FireEndTime", "Время окончания пожара", , , "Prop.FireEndTime", True, False
    AddNewElement "CurrentTime", "Текущее астрономическое время схемы", , , , True, True
    AddNewElement "CurrentTimeDelta", "Текущее относительное время схемы", "Ч+{$}", , , True, False
    
End Sub

Private Function AddNewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_resultStrPattern As String = "", _
                        Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InTacticDataForm As Boolean = True) As Element
'Создает новый элемент, добавляет его в коллекцию элементов и возвращает на него ссылку
    Set AddNewElement = AddElement(NewElement(a_ID, a_CallName, a_resultStrPattern, a_startValue, a_CellID, a_InShapes, a_InTacticDataForm))
End Function

Private Function NewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_resultStrPattern As String = "", _
                        Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InTacticDataForm As Boolean = True) As Element
'Создает новый элемент, задает ему свойства и возвращает на него ссылку
Dim elem As Element

    Set elem = New Element
    Set NewElement = elem.SetShell(Me).Activate(a_ID, a_CallName, a_resultStrPattern, a_startValue, a_CellID, a_InShapes, a_InTacticDataForm)
End Function

Private Function AddElement(ByRef elem As Element) As Element
'Добавляет элемент в коллекцию элементов и возвращает на него ссылку
Dim nextElem As Element
    
    Set nextElem = FindHigherElement(elem)
    'Основная коллекция
    If nextElem Is Nothing Then
        elemCollection.Add elem, elem.ID
    Else
        elemCollection.Add elem, elem.ID, nextElem.ID
    End If
    
    'Коллекция элементов с расчетными ячейками
'    Debug.Assert elem.cellID = ""
'    Debug.Assert elem.cellID <> "Prop.FireTime"
    If elem.cellID <> "" Then elemWithCellsCollection.Add elem, elem.cellID
    
    Set AddElement = elem
End Function

Private Function FindHigherElement(ByRef elemIn As Element) As Element
'Возвращает элемент в коллекции elemCollection с именем вызова больше чем у текущего (тот элемент перед которым нужно будет вставить новый)
Dim elem As Element
    
'    For Each elem In elemCollection
'        If elem.callName > elemIn.callName Then
'            Set FindHigherElement = elem
'            Exit Function
'        End If
'    Next elem
End Function

'-------------------------------------Публичные процедуры классса--------------------------------------------
Public Sub PrintState(Optional ByVal elemIDs As String = "")
'Печатает в Debug все элементы и их значения
Dim elem As Element
Dim elemArray() As String
Dim i As Integer
    
    If elemIDs = "" Then
        For Each elem In elemCollection
            elem.PrintState
        Next elem
    Else
        elemArray = Split(elemIDs, ";")
        For i = 0 To UBound(elemArray)
            Set elem = ByID(elemArray(i))
            If Not elem Is Nothing Then elem.PrintState
        Next i
    End If
    
End Sub


'-------------------------------------Публичные функции классса----------------------------------------------
Public Function ByID(ByVal a_elemID As String) As Element
'Функция возвращает ссылку на Элемент с ID = a_elemID
    On Error GoTo ex
    Set ByID = elemCollection.item(a_elemID)
    
Exit Function
ex:
    Set ByID = Nothing
End Function

Public Function ByCellID(ByVal a_elemCellID As String) As Element
'Функция возвращает ссылку на Элемент с именем ячейки CellID = a_elemCellID
    On Error GoTo ex
    Set ByCellID = elemWithCellsCollection.item(a_elemCellID)
    
Exit Function
ex:
    Set ByCellID = Nothing
End Function

Public Function ByCallName(ByVal a_elemCallName As String) As Element
'Функция возвращает ссылку на элемент с Именем вызова callName = a_elemCallName
Dim elem As Element

    For Each elem In elemCollection
        If elem.callName = a_elemCallName Then
            Set ByCallName = elem
            Exit Function
        End If
    Next elem
'Если ничего не найдено возвращаем Nothing
Set ByCallName = Nothing
End Function

Public Function GetElementsCollection(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False) As Collection
'Функция возвращает коллекцию элементов отобранных по поисковой строке
Dim coll As Collection
Dim elem As Element
Dim addFlag As Boolean
    
    Set coll = New Collection
    
    'перебираем все элементы и проверяем соответствие
    For Each elem In elemCollection
        addFlag = False
        If InStr(1, elem.ID, searchString, vbTextCompare) > 0 Then
            addFlag = True
        End If
        
        If searchByCallName Then
            If InStr(1, elem.callName, searchString) > 0 Then
                addFlag = True
            End If
        End If
        
        If addFlag Then
            coll.Add elem
        End If
    Next elem
    
Set GetElementsCollection = coll
End Function

Public Function Clear() As ElementsShell
'Обнуляем значения свойств и при водим значения элементов в начальное состояние
Dim elem As Element

    'перебираем все элементы и обнуляем каждый из них
    For Each elem In elemCollection
        elem.Clear
    Next elem
    
    Set Clear = Me
End Function

Public Function RaiseByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String, Optional ByVal dataType As VisUnitCodes = visNumber) As ElementsShell
'Увеличиваем значение элемента по имени ячейки
'a_CellIDs - перечисление через строчку с запятой имен ячеек в виде "User.IndexPers"
Dim elem As Element
Dim cellIDArray() As String
Dim i As Integer
    
    cellIDArray = Split(a_CellIDs, ";")
    
    For i = 0 To UBound(cellIDArray)
'Вариант с уникальным именем ячейки для каждого элемента - так быстрее, но нельзя считать для многих элементов
        Set elem = ByCellID(cellIDArray(i))
        If Not elem Is Nothing Then
'            Select Case dataType
'                Case Is = visNumber
'                    elem.RaiseValue CellVal(shp, cellIDArray(i), dataType)
'                Case Is = visUnitsString
'                    elem.SetValue CellVal(shp, cellIDArray(i), dataType)
'                Case Is = visDate
'                    elem.RaiseValue CellVal(shp, cellIDArray(i), dataType)
'            End Select
            elem.RaiseValue cellVal(shp, cellIDArray(i), dataType)
        End If
    Next i
    
Set RaiseByCellID = Me
End Function

Public Function SetValByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String, Optional ByVal dataType As VisUnitCodes = visNumber) As ElementsShell
'Устанавливаем значение элемента по имени ячейки
'a_CellIDs - перечисление через строчку с запятой имен ячеек в виде "User.IndexPers"
Dim elem As Element
Dim cellIDArray() As String
Dim i As Integer
    
    cellIDArray = Split(a_CellIDs, ";")
    
    For i = 0 To UBound(cellIDArray)
'Вариант с уникальным именем ячейки для каждого элемента - так быстрее, но нельзя считать для многих элементов
        Set elem = ByCellID(cellIDArray(i))
        If Not elem Is Nothing Then
            elem.SetValue cellVal(shp, cellIDArray(i), dataType)
        End If
    Next i
    
Set SetValByCellID = Me
    
End Function

Public Function Raise(ByVal elemIDs As String, Optional ByVal val As Variant = 1) As ElementsShell
'Увеличиваем значение указанных ячеек. Список указывается через точку с запятой
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!Внимание! Элемент с ID=" & elemArray(i) & " отсутствет!"
        Else
            elem.RaiseValue val
        End If
    Next i
Set Raise = Me
End Function

Public Function SetVal(ByVal elemIDs As String, Optional ByVal val As Variant = 1) As ElementsShell
'Устанавливаем значения указанных ячеек. Список указывается через точку с запятой
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!Внимание! Элемент с ID=" & elemArray(i) & " отсутствет!"
        Else
            elem.SetValue val
        End If
    Next i
Set SetVal = Me
End Function

Public Function CallNames(Optional ByVal delimiter As String = ";") As String
'Функция возвращает строку имен элементов с указанным разделителем
Dim elem As Element
Dim tmpStr As String
    
    If elemCollection.count = 0 Then
        CallNames = ""
        Exit Function
    End If
    
    For Each elem In elemCollection
        If elem.inShapes Then
            tmpStr = tmpStr & elem.callName & delimiter
        End If
    Next elem
    
    If Len(tmpStr) = 0 Then
        CallNames = ""
        Exit Function
    End If

CallNames = Left(tmpStr, Len(tmpStr) - Len(delimiter))
End Function


'------Функции возвращающие значения элементов-----------------
Public Function Result(ByVal elemID As String) As Variant
'Получаем значение элемента с указанным elemID
Dim elem As Element
    Set elem = ByID(elemID)
    If Not elem Is Nothing Then
        Result = elem.Result
    Else
        Result = 0
    End If
End Function
Public Function ResultByCN(ByVal callName As String) As Variant
'Получаем значение элемента с указанным CallName
Dim elem As Element
    
    Set elem = ByCallName(callName)
    
    If Not elem Is Nothing Then
        ResultByCN = elem.Result
    Else
        ResultByCN = 0
    End If
End Function

Public Function Sum(ByVal elemIDs As String) As Variant
'Получаем сумму значений элементов с указанными через точку с запятой именами элементов - elemIDs
Dim elemArray() As String
Dim elem As Element
Dim i As Integer
Dim tmp As Single
    
    elemArray = Split(elemIDs, ";")
        
    For i = 0 To UBound(elemArray)
        Set elem = ByID(elemArray(i))
        If elem Is Nothing Then
            Debug.Print "!Внимание! Элемент с ID=" & elemArray(i) & " отсутствет!"
        Else
            tmp = tmp + elem.Result
        End If
    Next i
Sum = tmp
End Function



