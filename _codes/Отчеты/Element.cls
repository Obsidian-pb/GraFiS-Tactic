VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'================Класс элемента оперативно-тактической обстановки для расчета=====================================================================================================================
Private p_ID As String                               'Переменная для хранения ID объекта
Private p_callName As String                         'Имя вызова (для использования в фигурах)
Private p_value As Variant                          'Текущее значение калькулятора
Private p_startValue As Variant                     'Стартовое значение калькулятора
Private p_state As Boolean                          'Состояние объекта (активирован или нет)

'--------------------Свойства-------------------------
Public Property Get ID() As String
'Идентификатор калькулятора
    ID = p_ID
End Property
Public Property Get dataType() As String
'Тип данных используемый калькулятором
    dataType = TypeName(p_value)
End Property
Public Property Get callName() As String
'Имя вызова
    callName = p_callName
End Property
Public Property Get Result() As Variant
'Текущее значение калькулятора
    Result = p_value
End Property


'--------------------Публичные функции класса--------------------
Public Function Activate(ByVal a_ID As String, ByVal a_CallName As String, ByVal a_startValue As Variant) As Element
'Активируем объект
    p_ID = a_ID
    p_callName = a_CallName
    p_startValue = a_startValue
    p_state = True
End Function

Public Function RaiseValue(Optional ByVal diffValue As Variant = 1) As Element
'Увеличиваем текущее значение на diffValue
    p_value = p_value + diffValue
    Set RaiseValue = Me
End Function

Public Function SetValue(ByVal newValue As Variant) As Element
'Устанавливаем новое значение
    p_value = newValue
    Set SetValue = Me
End Function

Public Function Clear() As Element
'Обнуляем калькулятор
    p_value = p_startValue
    Set Clear = Me
End Function

'--------------------Публичные процедуры




