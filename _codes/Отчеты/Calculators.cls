VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================Клас-обертка для коллекции калькуляторов==========================================
Private calcCollection As Collection



'-------------------------------------Базовые процедуры классса------------------------------------
Private Sub Class_Initialize()
    Set calcCollection = New Collection
    fillCalcCollection
End Sub
Private Sub Class_Terminate()
    Set calcCollection = Nothing
End Sub

'-------------------------------------Свойства класса----------------------------------------------
Public Property Get Count() As Integer
'Количество калькуляторов в коллекции
    Count = calcCollection.Count
End Property




'-------------------------------------Внутренние процедуры и функции классса----------------------------------------------
'!!!Заполнение коллекции калькуляторов
Private Sub fillCalcCollection()
'Заполняем коллекцию калькуляторов
Dim calc As ICalculator
    Set calc = New xCalcTest
        calcCollection.Add calc, calc.ID
    Set calc = New xCalcGDZS
        calcCollection.Add calc, calc.ID
    
    
    
    
    
End Sub

'-------------------------------------Публичные процедуры классса--------------------------------------------
Public Sub Calculate(ByRef shp As Visio.Shape)
'!!!Основная процедура запускающая цикл расчета всех калькуляторов для данной фигуры
Dim calc As ICalculator

    For Each calc In calcCollection
        calc.Calculate shp
    Next calc
End Sub

'-------------------------------------Публичные функции классса----------------------------------------------
Public Function GetCalculatorByID(ByVal a_calcID As String) As ICalculator
'Функция возвращает ссылку на калькулятор с ID = a_calcID
    On Error GoTo EX
    Set GetCalculatorByID = calcCollection.Item(a_calcID)
    
Exit Function
EX:
    Set GetCalculatorByID = Nothing
End Function

Public Function GetCalculatorByCallName(ByVal a_calcCallName As String) As ICalculator
'Функция возвращает ссылку на калькулятор с Именем вызова callName = a_calcID
Dim calc As ICalculator
    
    For Each calc In calcCollection
        If calc.callName = a_calcCallName Then
            Set GetCalculatorByCallName = calc
            Exit Function
        End If
    Next calc
'Если ничего не найдено возвращаем Nothing
Set GetCalculatorByCallName = Nothing
End Function

Public Function CallNamesArray(ByVal delimiter As String) As String
'Функция возвращает строку имен калькуляторов с указанным разделителем
Dim calc As ICalculator
Dim tmpStr As String
    
    If calcCollection.Count = 0 Then
        CallNamesArray = ""
        Exit Function
    End If
    
    For Each calc In calcCollection
        tmpStr = tmpStr & calc.callName & delimiter
    Next calc
    
CallNamesArray = Left(tmpStr, Len(tmpStr) - Len(delimiter))
End Function







