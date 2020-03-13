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
Public bo_GDZSRezRoundUp As Boolean 'Округление рез.звеньев в большую сторону


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
    AddNewElement "FireSquare", "Площадь пожара"
    AddNewElement "ExtSquare", "Площадь тушения"
    AddNewElement "FactStreamW", "Расход воды фактический"
    AddNewElement "IntenseWater", "Интенсивность подачи воды"
    AddNewElement "NeedStreamW", "Расход воды требуемый"
    
    AddNewElement "PersonnelHave", "Имеется личного состава"
    AddNewElement "PersonnelNeed", "Требуется личного состава"
    
    AddNewElement "StvolWHave", "Водяных стволов подано"
    AddNewElement "StvolWBHave", "Водяных стволов Б подано"
    AddNewElement "StvolWAHave", "Водяных стволов А подано"
    AddNewElement "StvolWLHave", "Водяных лафетных стволов подано"
    
    AddNewElement "StvolFoamHave", "Пенных стволов подано"
    AddNewElement "StvolGasHave", "Газовых стволов подано"
    AddNewElement "StvolPowderHave", "Порошковых стволов подано"
    
    AddNewElement "MainPAHave", "Основных ПА общего применения имеется"
    AddNewElement "TargetedPAHave", "Основных ПА целевого применения имеется"
    AddNewElement "MainOverallCount", "Основных ПА имеется"
    AddNewElement "SpecialPAHave", "Специальных ПА имеется"
    AddNewElement "OtherTechincsHave", "Прочей техники имеется"
    
    AddNewElement "GetedWaterValue", "Забирается воды из внешних источников"
    AddNewElement "GetedWaterValueMax", "Забирается воды из внешних источников Максимум"
    AddNewElement "GetingWaterCount", "Устройств забирающих воду (кол-во)"
    
    AddNewElement "GDZSChainsCountHave", "Звеньев ГДЗС работает"
    AddNewElement "GDZSChainsCountNeed", "Звеньев ГДЗС требуется"
    AddNewElement "GDZSChainsRezCountHave", "Звеньев ГДЗС в резерве имеется"
    AddNewElement "GDZSChainsRezCountNeed", "Звеньев ГДЗС в резерве требуется"
    AddNewElement "GDZSMansCountHave", "Газодымозащитников работает"
    AddNewElement "GDZSMansCountNeed", "Газодымозащитников требуется"
    AddNewElement "GDZSMansRezCountHave", "Газодымозащитников в резерве имеется"
    AddNewElement "GDZSMansRezCountNeed", "Газодымозащитников в резерве требуется"
    'AddNewElement "GDZSDiscr", "Несоотвествие количества газодымозащитников в звене сложным условиям", 0
    AddNewElement "GDZSPBCount", "Количество ПБ ГДЗС"
    AddNewElement "GDZSKPPCount", "Количество КПП ГДЗС"
    
    AddNewElement "Hoses38Count", "Рукавов 38мм используется"
    AddNewElement "Hoses51Count", "Рукавов 51мм используется"
    AddNewElement "Hoses66Count", "Рукавов 66мм используется"
    AddNewElement "Hoses77Count", "Рукавов 77мм используется"
    AddNewElement "Hoses89Count", "Рукавов 89мм используется"
    AddNewElement "Hoses110Count", "Рукавов 110мм используется"
    AddNewElement "Hoses150Count", "Рукавов 150мм используется"
    AddNewElement "Hoses200Count", "Рукавов 200мм используется"
    AddNewElement "Hoses250Count", "Рукавов 250мм используется"
    AddNewElement "Hoses300Count", "Рукавов 300мм используется"
    AddNewElement "HosesCount", "Рукавов используется всего"
    AddNewElement "HosesLength", "Длина рукавных линий"
    
    AddNewElement "Hoses38Have", "Рукавов 38мм имеется", , "Prop.Hose38"
    AddNewElement "Hoses51Have", "Рукавов 51мм имеется", , "Prop.Hose51"
    AddNewElement "Hoses66Have", "Рукавов 66мм имеется", , "Prop.Hose66"
    AddNewElement "Hoses77Have", "Рукавов 77мм имеется", , "Prop.Hose77"
    AddNewElement "Hoses89Have", "Рукавов 89мм имеется", , "Prop.Hose89"
    AddNewElement "Hoses110Have", "Рукавов 110мм имеется", , "Prop.Hose110"
    AddNewElement "Hoses150Have", "Рукавов 150мм имеется", , "Prop.Hose150"
    AddNewElement "Hoses200Have", "Рукавов 200мм имеется", , "Prop.Hose200"
    AddNewElement "Hoses250Have", "Рукавов 250мм имеется", , "Prop.Hose250"
    AddNewElement "Hoses300Have", "Рукавов 300мм имеется", , "Prop.Hose300"
    AddNewElement "HosesHave", "Рукавов имеется всего"
    AddNewElement "HoseSkatka", "Напорных рукавах в скатках"
    
    AddNewElement "WaterValueNeed", "Воды требуется"
    AddNewElement "WaterValueHave", "Воды имеется"
    
    AddNewElement "ACCount", "АЦ имеется"
    AddNewElement "AGCount", "АГ имеется"
    AddNewElement "ALCount", "АЛ имеется"
    AddNewElement "AKPCount", "АКП имеется"
    AddNewElement "MVDCount", "Техники МВД имеется"
    AddNewElement "SMPCount", "Техники СМП имеется"
    AddNewElement "TotalTechCount", "Всего техники имеется"
    AddNewElement "TotalFireCount", "Всего пожарной техники имеется"
    
    AddNewElement "BUCount", "БУ создано"
    AddNewElement "SPRCount", "СПР создано"
    AddNewElement "ShtabCount", "Количество штабов", 0, , False, False
    AddNewElement "RNBDCount", "Количество решющих направления боевых действий", 0, , False, False
    AddNewElement "OchagCount", "Количество очагов", 0, , False, False
    AddNewElement "SpreadCount", "Количество направлдений распространения огня", 0, , False, False
    AddNewElement "SmokeCount", "Задымление", 0, , False, False
    AddNewElement "WaterSourceCount", "Водоисточники", 0, , False, False
    AddNewElement "DistanceCount", "Расстояния", 0, , False, False
    AddNewElement "LinesCount", "Отдельные рукавные линии", 0, , False, False
    AddNewElement "LinesPosCount", "Подписи положения рукавных линий", 0, , False, False
    AddNewElement "LinesLableCount", "Подписи рукавных линий", 0, , False, False
    AddNewElement "WorklinesCount", "Количество рабочих линий", 0, , False, False
    AddNewElement "BuildCount", "Количество зданий", 0, , False, False
    AddNewElement "SOCount", "Количество степеней огнестойкости", 0, , False, False
    AddNewElement "FireCount", "Количество площадей горения", 0, , False, False
    AddNewElement "OrientCount", "Ориентиры", 0, , False, False
    
    AddNewElement "AllHosesWithLabels", "Все рукавные линии с подписями", True, , False, False
    AddNewElement "AllHosesWithPos", "Все рабочие рукавные линии с подписями положения", True, , False, False
    
    
    
    
    
    '*Сделать инструментальную процедуру, которая будет искать Элемент по тексту и выдавать возможные варианты
    
    

    
End Sub

Private Function AddNewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InMaster As Boolean = True) As Element
'Создает новый элемент, добавляет его в коллекцию элементов и возвращает на него ссылку
    Set AddNewElement = AddElement(NewElement(a_ID, a_CallName, a_startValue, a_CellID, a_InShapes, a_InMaster))
End Function

Private Function NewElement(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InMaster As Boolean = True) As Element
'Создает новый элемент, задает ему свойства и возвращает на него ссылку
Dim elem As Element

    Set elem = New Element
    Set NewElement = elem.Activate(a_ID, a_CallName, a_startValue, a_CellID, a_InShapes, a_InMaster)
End Function

Private Function AddElement(ByRef elem As Element) As Element
'Добавляет элемент в коллекцию элементов и возвращает на него ссылку
    'Основная коллекция
    elemCollection.Add elem, elem.ID
    'Коллекция элементов с расчетными ячейками
    If elem.cellID <> "" Then elemWithCellsCollection.Add elem, elem.cellID
    
    Set AddElement = elem
End Function

'-------------------------------------Публичные процедуры классса--------------------------------------------



'-------------------------------------Публичные функции классса----------------------------------------------
Public Function ByID(ByVal a_elemID As String) As Element
'Функция возвращает ссылку на Элемент с ID = a_elemID
    On Error GoTo EX
    Set ByID = elemCollection.Item(a_elemID)
    
Exit Function
EX:
    Set ByID = Nothing
End Function

'Public Function ByCallName(ByVal a_elemCallName As String) As Element
''Функция возвращает ссылку на элемент с Именем вызова callName = a_elemCallName
'Dim elem As Element
'
'    For Each elem In elemCollection
'        If elem.callName = a_elemCallName Then
'            Set ByCallName = elem
'            Exit Function
'        End If
'    Next elem
''Если ничего не найдено возвращаем Nothing
'Set ByCallName = Nothing
'End Function

Public Function Clear(ByRef shp As Visio.Shape) As ElementsShell
'Обнуляем значения свойств и при водим значения элементов в начальное состояние
Dim elem As Element

    'перебираем все элементы и обнуляем каждый из них
    For Each elem In elemCollection
        elem.Clear
    Next elem
    
    Set Clear = Me
End Function

Public Function CalculateByCellID(ByRef shp As Visio.Shape, ByVal a_CellID As String) As ElementsShell
'Вычисляем элемент по имени ячейки
Dim elem As Element
    
    On Error GoTo NoSuchElement
    Set elem = elemWithCellsCollection.Item(a_CellID)
    elem.RaiseValue (CellVal(shp, "User.FireSquare"))
    
    Set CalculateByCellID = Me
Exit Function
NoSuchElement:
    Set CalculateByCellID = Me
End Function

Public Function Raise(ByVal elemNames As String, Optional ByVal val As Variant = 1) As ElementsShell
'Увеличиваем значение указанных ячеек. Список указывается через точку с запятой
    Stop
End Function

Public Function SetVal(ByVal elemNames As String, Optional ByVal val As Variant = 1) As ElementsShell
'Устанавливаем значения указанных ячеек. Список указывается через точку с запятой
    Stop
End Function

Public Function CallNames(ByVal delimiter As String) As String
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








