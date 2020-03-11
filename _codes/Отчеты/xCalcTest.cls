VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xCalcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICalculator

'================Класс калькулятора элемента оперативно-тактической обстановки================
Const p_ID = "test"                                         'Идентификатор калькулятора
Const p_dataType = CalculatorResultTypes.crt_Number         'Тип данных используемый калькулятором
Const p_callName = "Тестовый калькулятор"                   'Именем вызова (для использования в фигурах)

Private p_value As Integer                                  'Текущее значение калькулятора


'----------------Неизменяемые при наследовании свойства--------------------------------------
Private Property Get ICalculator_callName() As String
'Имя вызова калькулятора
    ICalculator_callName = p_callName
End Property

Private Property Get ICalculator_dataType() As CalculatorResultTypes
'Тип данных калькулятора
    ICalculator_dataType = p_dataType
End Property

Private Property Get ICalculator_ID() As String
'Идентификатор калькулятора
    ICalculator_ID = p_ID
End Property

Private Property Get ICalculator_Result() As Variant
'Текущее значение
    ICalculator_Result = p_value
End Property





Private Sub ICalculator_RaiseValue(Optional ByVal diffValue As Variant = 1)
    p_value = p_value + diffValue
End Sub

Private Sub ICalculator_Calculate(shp As Visio.Shape)
'Пересчетываем калькулятор и передаем ему ссылку на рассматриваемую фигуру
    If shp.CellExists("User.IndexPers", 0) Then
        ICalculator_RaiseValue
    End If
End Sub
