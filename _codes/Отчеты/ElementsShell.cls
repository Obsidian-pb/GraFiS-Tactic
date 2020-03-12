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

'================Клас-обертка для коллекции элементов==========================================
Private elemCollection As Collection



'-------------------------------------Базовые процедуры классса------------------------------------
Private Sub Class_Initialize()
    Set elemCollection = New Collection
    fillElemCollection
End Sub
Private Sub Class_Terminate()
    Set elemCollection = Nothing
End Sub

'-------------------------------------Свойства класса----------------------------------------------
Public Property Get Count() As Integer
'Количество элементов в коллекции
    Count = elemCollection.Count
End Property




'-------------------------------------Внутренние процедуры и функции классса----------------------------------------------
Private Sub fillElemCollection()
'Заполняем коллекцию элементов
    AddNewElement "FireSquare", "Площадь пожара", 0
    AddNewElement "ExtSquare", "Площадь тушения", 0
    AddNewElement "FactStreamW", "Фактический расход воды", 0
    AddNewElement "IntenseWater", "Интенсивность подачи воды", 0
    AddNewElement "NeedStreamW", "Требуемый расход воды", 0
    AddNewElement "HasLabels", "Проба Для Bool", True
    
End Sub

Private Function AddNewElement(ByVal a_ID As String, ByVal a_CallName As String, ByVal a_startValue As Variant) As Element
'Создает новый элемент, задает ему свойства, добавляет его в коллекцию элементов и возвращает на него ссылку
Dim elem As Element
    
    Set elem = New Element
    Set AddNewElement = elem.Activate(a_ID, a_CallName, a_startValue)
    elemCollection.Add elem, elem.ID
    
End Function

'-------------------------------------Публичные процедуры классса--------------------------------------------
Public Sub Clear(ByRef shp As Visio.Shape)
'Обнуляем значения свойств и при водим значения элементов в начальное состояние
Dim elem As Element

    'перебираем все элементы и обнуляем каждый из них
    For Each elem In elemCollection
        elem.Clear
    Next elem

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

Public Function CallNames(ByVal delimiter As String) As String
'Функция возвращает строку имен элементов с указанным разделителем
Dim elem As Element
Dim tmpStr As String
    
    If elemCollection.Count = 0 Then
        CallNames = ""
        Exit Function
    End If
    
    For Each elem In elemCollection
        tmpStr = tmpStr & elem.callName & delimiter
    Next elem
    
CallNames = Left(tmpStr, Len(tmpStr) - Len(delimiter))
End Function








