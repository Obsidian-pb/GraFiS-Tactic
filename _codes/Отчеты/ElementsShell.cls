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
Public Property Get Count() As Integer
'Количество элементов в коллекции
    Count = elemCollection.Count
End Property




'-------------------------------------Внутренние процедуры и функции классса----------------------------------------------
Private Sub fillElemCollection()
'Заполняем коллекцию элементов

    AddNewElement "RSCHSTechincsHave", "Техники РСЧС имеется", "{$} ({TotalFireCount} - пожарной и АС)"
    AddNewElement "MainPAHave", "ПА общего применения имеется", , , , , 0
    AddNewElement "TargetedPAHave", "ПА целевого применения имеется", , , , , 0
    AddNewElement "MainOverallHave", "ПА основных имеется", "{$} ({TargetedPAHave} - цел.прим., {MainPAHave} - общ.прим.)"
    AddNewElement "PAHighHave", "ПА высотных имеется", , , , , 0
    AddNewElement "SpecialPAHave", "ПА специальных имеется", "{$} ({PAHighHave} - высотных)"
    AddNewElement "OtherTechincsHave", "Прочей пожарной и аварийно-спасательной техники имеется"
    
    AddNewElement "ACCount", "АЦ имеется", , , , , 0
    AddNewElement "ACNeed", "АЦ требуется", , , , , 0
    AddNewElement "ANRNeed", "АНР требуется", , , , , 0
    AddNewElement "AGCount", "АГДЗС имеется", , , , , 0
    AddNewElement "ALCount", "АЛ имеется", , , , , 0
    AddNewElement "AKPCount", "АКП имеется", , , , , 0
    AddNewElement "MVDCount", "Техники МВД имеется", , , , , 0
    AddNewElement "SMPCount", "Техники СМП имеется", , , , , 0
    AddNewElement "TechnicsNotMchsHave", "Техники прочих служб имеется", "{$} ({MVDCount} - МВД, {SMPCount} - СМП)"
    AddNewElement "TechnicsNotMchsOtherHave", "Прочей техники имеется", , , , , 0
    AddNewElement "TotalTechCount", "Техники всего имеется", , , , , 0
    AddNewElement "TotalFireCount", "Пожарной и аварийно-спасательной техники всего имеется", , , , , 0
    
    AddNewElement "FireSquare", "Площадь пожара", "{$} м.кв. ({ExtSquare} м.кв. - площадь тушения)", , "User.FireSquare"
    AddNewElement "ExtSquare", "Площадь тушения", , , "User.ExtSquare", , 0
    AddNewElement "FactStreamW", "Расход воды фактический", "{$} л/с ({NeedStreamW} - л/с требуемый)", , "User.PodOut"
    AddNewElement "IntenseWater", "Интенсивность подачи воды"
    AddNewElement "NeedStreamW", "Расход воды требуемый", "{$} л/с", , "User.NeedStreamW", , 0
    
    AddNewElement "PersonnelHave", "Личного состава имеется", "{$} чел. ({PersonnelNeed} чел. - требуется)"
    AddNewElement "PersonnelNeed", "Личного состава требуется", , , "Prop.Personnel", , 0
    
    AddNewElement "StvolWHave", "Стволов водяных подано", "{$} ({StvolWBHave} - ств. 'Б', {StvolWAHave} - ств. 'А', {StvolWLHave} - ств. лаф.)"
    AddNewElement "StvolWBHave", "Стволов Б подано", , , , , 0
    AddNewElement "StvolWAHave", "Стволов А подано", , , , , 0
    AddNewElement "StvolWLHave", "Стволов водяных лафетных подано", , , , , 0
    AddNewElement "StvolWBNeed", "Стволов Б требуется", , , , , 0
    AddNewElement "StvolWANeed", "Стволов А требуется", , , , , 0
    AddNewElement "StvolWLNeed", "Стволов водяных лафетных требуется", , , , , 0
    
    AddNewElement "StvolFoamHave", "Стволов пенных подано"
    AddNewElement "StvolGasHave", "Стволов газовых подано"
    AddNewElement "StvolPowderHave", "Стволов порошковых подано"

    AddNewElement "GetedWaterValue", "Забирается воды из внешних источников", "{$} л/с (max = {GetedWaterValueMax} л/с )", , , , False
    AddNewElement "GetedWaterValueMax", "Забирается воды из внешних источников Максимум", , , , , False
    AddNewElement "GetingWaterCount", "Установлено на водоисточники", "{$} ({PANeedOnWaterSource} - требуется установить)"
    AddNewElement "PANeedOnWaterSource", "Требуется установить на водоисточники", , , , , 0
    AddNewElement "WaterEternal", "Бесконечный запас воды", , False, , False, False

    AddNewElement "GDZSChainsCountWork", "Звеньев ГДЗС работает", "{$} ({GDZSMansCountWork} газодымозащитников)"
    AddNewElement "GDZSChainsCountWorkSoft", "Звеньев ГДЗС работает в стандартных условиях", , , , , False
    AddNewElement "GDZSChainsCountWorkHard", "Звеньев ГДЗС работает в сложных условиях", , , , , False
    AddNewElement "GDZSChainsCountNeed", "Звеньев ГДЗС требуется", , , , , 0
    AddNewElement "GDZSChainsRezCountHave", "Звеньев ГДЗС в резерве имеется", "{$} ({GDZSMansRezCountHave} газодымозащитников)", , , , 0
    AddNewElement "GDZSChainsRezCountNeed", "Звеньев ГДЗС в резерве требуется", , , , , 0
'    AddNewElement "GDZSChainsCanBeFormed", "Звеньев ГДЗС Может быть сформировано"
    AddNewElement "GDZSMansCountWork", "Газодымозащитников работает", , , , , 0
    AddNewElement "GDZSMansCountNeed", "Газодымозащитников требуется", , , , , 0
    AddNewElement "GDZSMansRezCountHave", "Газодымозащитников в резервных звеньях имеется", , , , , 0
    AddNewElement "GDZSMansRezCountNeed", "Газодымозащитников в резерве требуется", , , , , 0
    AddNewElement "GDZSDiscr", "Несоотвествие количества газодымозащитников в звене сложным условиям", , False
    AddNewElement "GDZSPBCount", "ПБ ГДЗС имеется", , , , , 0
    AddNewElement "GDZSKPPCount", "КПП ГДЗС имеется", , , , , 0

    AddNewElement "Hoses38Count", "Рукавов 38мм используется", , , , , 0
    AddNewElement "Hoses51Count", "Рукавов 51мм используется", , , , , 0
    AddNewElement "Hoses66Count", "Рукавов 66мм используется", , , , , 0
    AddNewElement "Hoses77Count", "Рукавов 77мм используется", , , , , 0
    AddNewElement "Hoses89Count", "Рукавов 89мм используется", , , , , 0
    AddNewElement "Hoses110Count", "Рукавов 110мм используется", , , , , 0
    AddNewElement "Hoses150Count", "Рукавов 150мм используется", , , , , 0
    AddNewElement "Hoses200Count", "Рукавов 200мм используется", , , , , 0
    AddNewElement "Hoses250Count", "Рукавов 250мм используется", , , , , 0
    AddNewElement "Hoses300Count", "Рукавов 300мм используется", , , , , 0
    AddNewElement "HosesCount", "Рукавов используется всего", "{$} (" & _
                                "{Hoses38Count} - 38мм, {Hoses51Count} - 51мм, {Hoses66Count} - 66мм, {Hoses77Count} - 77мм, {Hoses89Count} - 89мм," & _
                                "{Hoses110Count} - 110мм, {Hoses150Count} - 150мм, {Hoses200Count} - 200мм, {Hoses250Count} - 250мм, {Hoses300Count} - 300мм" & _
                                ")"
    AddNewElement "HosesLength", "Рукавных линий длина", "{$} м"

    AddNewElement "Hoses38Have", "Рукавов 38мм имеется", , , "Prop.Hose38", , 0
    AddNewElement "Hoses51Have", "Рукавов 51мм имеется", , , "Prop.Hose51", , 0
    AddNewElement "Hoses66Have", "Рукавов 66мм имеется", , , "Prop.Hose66", , 0
    AddNewElement "Hoses77Have", "Рукавов 77мм имеется", , , "Prop.Hose77", , 0
    AddNewElement "Hoses89Have", "Рукавов 89мм имеется", , , "Prop.Hose89", , 0
    AddNewElement "Hoses110Have", "Рукавов 110мм имеется", , , "Prop.Hose110", , 0
    AddNewElement "Hoses150Have", "Рукавов 150мм имеется", , , "Prop.Hose150", , 0
    AddNewElement "Hoses200Have", "Рукавов 200мм имеется", , , "Prop.Hose200", , 0
    AddNewElement "Hoses250Have", "Рукавов 250мм имеется", , , "Prop.Hose250", , 0
    AddNewElement "Hoses300Have", "Рукавов 300мм имеется", , , "Prop.Hose300", , 0
    AddNewElement "HosesHave", "Рукавов имеется всего", , , , , 0
    AddNewElement "HoseSkatka", "Напорных рукавов в скатках", , , , , 0
    
    AddNewElement "WaterValueNeed10min", "Воды требуется (на 10мин)", "{$} л.", , , , 0
    AddNewElement "WaterValueHave", "Воды имеется", "{$} л."
    AddNewElement "FoamValueeHave", "Пенообразователя имеется", "{$} л.", , "Prop.Foam"

    AddNewElement "BUCount", "БУ создано"
    AddNewElement "SPRCount", "СПР создано"
    AddNewElement "ShtabCount", "Количество штабов", , 0, , False, False
    AddNewElement "RNBDCount", "Количество решающих направления боевых действий", , 0, , False, False
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
    AddNewElement "DeadCount", "Погибших", , , , , 0

    AddNewElement "AllHosesWithLabels", "Все рукавные линии с подписями", , True, , False, False
    AddNewElement "AllHosesWithPos", "Все рабочие рукавные линии с подписями положения", , True, , False, False

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
    If elem.cellID <> "" Then elemWithCellsCollection.Add elem, elem.cellID
    
    Set AddElement = elem
End Function

Private Function FindHigherElement(ByRef elemIn As Element) As Element
'Возвращает элемент в коллекции elemCollection с именем вызова больше чем у текущего (тот элемент перед которым нужно будет вставить новый)
Dim elem As Element
    
    For Each elem In elemCollection
        If elem.callName > elemIn.callName Then
            Set FindHigherElement = elem
            Exit Function
        End If
    Next elem
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
    On Error GoTo EX
    Set ByID = elemCollection.Item(a_elemID)
    
Exit Function
EX:
    Set ByID = Nothing
End Function

Public Function ByCellID(ByVal a_elemCellID As String) As Element
'Функция возвращает ссылку на Элемент с именем ячейки CellID = a_elemCellID
    On Error GoTo EX
    Set ByCellID = elemWithCellsCollection.Item(a_elemCellID)
    
Exit Function
EX:
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

Public Function RaiseByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String) As ElementsShell
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
            elem.RaiseValue CellVal(shp, cellIDArray(i))
        End If
    Next i
    
Set RaiseByCellID = Me
End Function

Public Function SetValByCellID(ByRef shp As Visio.Shape, ByVal a_CellIDs As String) As ElementsShell
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
            elem.SetValue CellVal(shp, cellIDArray(i))
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
    
    If elemCollection.Count = 0 Then
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



