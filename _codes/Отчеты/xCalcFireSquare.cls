VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xCalcFireSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICalculator
'================Класс калькулятора элемента оперативно-тактической обстановки================
'----------------1 Скопировать код имеющегося класса------------------------------------------
'----------------2 Откорректировать p_ID, p_dataType и p_callName-----------------------------
'----------------3 Откорректировать код для своих целей---------------------------------------
'----------------4 Создать экземпляр объекта этого класса в Calculators-----------------------
'----------------5 Добавить объект в коллекцию Calculators.calcCollection---------------------
Const p_ID = "FireSquare"                                           'Идентификатор калькулятора
Const p_dataType = CalculatorResultTypes.crt_Number                 'Тип данных используемый калькулятором
Const p_callName = "Площадь пожара"                                 'Именем вызова (для использования в фигурах)

'<<<Члены данного класса------------------------------------
    Private latestSquareTime As Date                                'Время самой последней фигуры площади пожара
'>>>Члены данного класса------------------------------------


'<<<Неизменяемые члены--------------------------------------
    Private p_value As Integer                              'Текущее значение калькулятора
    
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
    'Изменяем текущее значение
        p_value = p_value + diffValue
    End Sub
    Private Sub ICalculator_SetValue(ByVal newValue As Variant)
    'Устанавливаем новое значение
        p_value = newValue
    End Sub
    Private Sub ICalculator_Clear()
    'Обнуляем калькулятор
        p_value = 0
        latestSquareTime = 0
    End Sub
'>>>Неизменяемые члены--------------------------------------





'---------------Члены унаследованные от интерфейса ICalculator
Private Sub ICalculator_Calculate(shp As Visio.Shape)
'Пересчетываем калькулятор и передаем ему ссылку на рассматриваемую фигуру
    'Проверяем, является ли фигура фигурой площади
    If IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara) Then
        'Проверяем, является ли фигура самой последней по времени
        If shp.Cells("Prop.SquareTime").Result(visDate) > latestSquareTime Then
            p_value = shp.Cells("User.FireSquare")
            latestSquareTime = shp.Cells("Prop.SquareTime") 'Устанавливаем новое значение наибольшего времени учета для площади
        ElseIf shp.Cells("Prop.SquareTime").Result(visDate) = latestSquareTime Then
            p_value = p_value + shp.Cells("User.FireSquare")
        End If
    End If
End Sub


'---------------Собственные члены класса


