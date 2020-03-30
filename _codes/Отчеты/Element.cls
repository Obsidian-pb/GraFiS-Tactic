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
Private p_ID As String                              'Переменная для хранения ID объекта
Private p_callName As String                        'Имя вызова (для использования в фигурах)
Private p_cellID As String                          'Имя рассчетной ячейки
Private p_value As Variant                          'Текущее значение калькулятора
Private p_resultStrPattern As String                'Шаблон результирующей строки
Private p_startValue As Variant                     'Стартовое значение калькулятора
Private p_state As Boolean                          'Состояние объекта (активирован или нет)
Private p_InShapes As Boolean                       'Отражать в списке свойств для фигур
Private p_inTacticForm As Boolean                   'Отражать в мастере проверок
Private p_HidedInMCheck As Boolean                  'Элемент не отображается в окне мастера проверки схемы

Private p_elementsShell As ElementsShell            'Ссылка на оболочку коллекции элементов

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
Public Property Get cellID() As String
'Имя рассчетной ячейки
    cellID = p_cellID
End Property
Public Property Get inShapes() As String
'Отображается ли в списке свойств в фигурах
    inShapes = p_InShapes
End Property
Public Property Get inTacticForm() As String
'Отображается ли в списке свойств в мастере
    inTacticForm = p_inTacticForm
End Property
Public Property Get Result() As Variant
'Текущее значение элемента
    Result = p_value
End Property
Public Property Get ResultStr() As Variant
'Текущее значение элемента в виде строки полученной из шаблона. Если шаблон не задан, возвращаем просто значение
    If p_resultStrPattern = "" Then
        ResultStr = CStr(p_value)
    Else
        ResultStr = FillStr
    End If
End Property
Public Property Get HidedInMCheck() As Boolean
'Элемент не отображается в окне мастера проверки схемы
    HidedInMCheck = p_HidedInMCheck
End Property
Public Property Let HidedInMCheck(ByVal newVal As Boolean)
'Элемент не отображается в окне мастера проверки схемы
    p_HidedInMCheck = newVal
End Property


'--------------------Публичные функции класса--------------------
Public Function Activate(ByVal a_ID As String, ByVal a_CallName As String, Optional ByVal a_resultStrPattern As String = "", _
                        Optional ByVal a_startValue As Variant = 0, Optional a_CellID As String = "", _
                        Optional a_InShapes As Boolean = True, Optional a_InTacticDataForm As Boolean = True) As Element
'Активируем объект
    p_ID = a_ID
    p_callName = a_CallName
    p_resultStrPattern = a_resultStrPattern
    p_value = a_startValue
    p_startValue = a_startValue
    p_cellID = a_CellID
    p_InShapes = a_InShapes
    p_inTacticForm = a_InTacticDataForm
    
    p_state = True
    
    Set Activate = Me
End Function

Public Function SetShell(ByRef a_elementsShell As ElementsShell)
    Set p_elementsShell = a_elementsShell
    Set SetShell = Me
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

Public Function PrintState() As Element
'Печатаем в Debug Имя вызова и значение элемента
    Debug.Print p_callName & ": " & p_value
    Set PrintState = Me
End Function

Public Function Clear() As Element
'Обнуляем калькулятор
    p_value = p_startValue
    Set Clear = Me
End Function

'--------------------Внутренние процедуры
Private Function FillStr() As String
'Формируем на основе шаблона строки результирующую строку
Dim arr() As String
Dim i As Integer
Dim tmpStr As String
    
    arr = ParsePattern(p_resultStrPattern)
    
    tmpStr = p_resultStrPattern
    For i = 0 To UBound(arr)
        If arr(i) = "$" Then
            tmpStr = Replace(tmpStr, "{" & arr(i) & "}", p_value)
        Else
            tmpStr = Replace(tmpStr, "{" & arr(i) & "}", p_elementsShell.Result(arr(i)))
        End If
    Next i
    
FillStr = tmpStr
End Function

Private Function ParsePattern(ByVal str As String) As String()
'Возвращает массив имеющихся в шаблоне имен элементов
Dim char As String
Dim i As Integer
Dim j As Integer
Dim s As Integer
Dim tmpStr() As String
    
    j = 0
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If char = "{" Then
            s = i + 1
        ElseIf char = "}" Then
            ReDim Preserve tmpStr(j)
            tmpStr(j) = Mid(str, s, i - s)
            j = j + 1
        End If
    Next i
ParsePattern = tmpStr
End Function



