VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================Интерфейс калькулятора элемента оперативно-тактической обстановки================
'Const p_ID As String                             'Идентификатор калькулятора
'Const p_dataType As CalculatorResultTypes        'Тип данных используемый калькулятором
'Const p_callName As String                       'Имя вызова (для использования в фигурах)
'Private p_reportPattern As String                  'Шаблон строки отчета места вставки данных указываются - {n}, где n - номер значения
'
'Private p_state As Boolean                         'Состояние объекта (активирован или нет)
'Private p_value As Variant                         'Текущее значение калькулятора


Public Property Get ID() As String
'Идентификатор калькулятора
End Property
Public Property Get dataType() As CalculatorResultTypes
'Тип данных используемый калькулятором
End Property
Public Property Get callName() As String
'Имя вызова
End Property
Public Property Get Result() As Variant
'Текущее значение калькулятора
End Property




Public Sub RaiseValue(Optional ByVal diffValue As Variant = 1)
'Для чисел
End Sub

Public Sub Calculate(ByRef shp As Visio.Shape)
'Процедура Пересчета калькулятора и передаем ему ссылку на рассматриваемую фигуру
End Sub



