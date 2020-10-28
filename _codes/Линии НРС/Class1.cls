VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton   'Рабочая линия
Attribute ComBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  'Всасывающая линия
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  'Магистральные линии
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  'Нормализация
Attribute ComBut4.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("Превращения").Controls("Рукав")
    Set ComBut2 = Application.CommandBars("Превращения").Controls("Всасывающий рукав")
    Set ComBut3 = Application.CommandBars("Превращения").Controls("Магистральная линия")
    Set ComBut4 = Application.CommandBars("Превращения").Controls("Расчет НРС")

End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Обратить в рукавную линию"
    
'    On Error GoTo EX

'---Проверяем
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub

'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру линии
                MakeHoseLine 51, 0
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
    
'            '---Обращаем фигуру в фигуру зону горения
'                MorphToLake

Exit Sub
ex:
    SaveLog Err, "ComBut_Click"
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие клика по кнопке "Всасывающая линия"

    On Error GoTo ex

'---Проверяем
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub
    
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру всасывающей линии
                MakeVHoseLine
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
ex:
    SaveLog Err, "ComBut2_Click"
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие клика по кнопке "Магистральная линия"

    On Error GoTo ex
    
'---Проверяем
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub
    
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру магистральной линии
                MakeHoseLine 77, 1
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If

Exit Sub
ex:
    SaveLog Err, "ComBut3_Click"
End Sub



'--------------------------------Внутренние функции класса---------------------
Private Function GetSelectedButtonCaption() As String
Dim v_Ctrl As CommandBarControl

    On Error GoTo ex
'---Проверяем какая кнопка нажата и в зависимости от этого выполняем действие
    For Each v_Ctrl In Application.CommandBars("Превращения").Controls
        If v_Ctrl.State = msoButtonDown Then
            GetSelectedButtonCaption = v_Ctrl.Caption
            Exit Function
        End If
    Next v_Ctrl
GetSelectedButtonCaption = ""

Exit Function
ex:
End Function


'-------------------------------Внешние методы класса---------------------
'Public Sub MakeHoseLine()
''Метод обращения в рукавную линию
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
''    On Error GoTo Tail
'
''---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---Идентифицируем активную фигуру
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---Получаем списки для фигуры
'    '---Запускаем процедуру получения списка Подразделений
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
'    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
'    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
'    HoseDiametersListImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
'    HoseResistanceValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
'    HoseMaxFlowValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
'    HoseWeightValueImport (ShpInd)
'
''---Устанавливаем значение текущего времени бкз ссылки
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---Отмечаем кнопку
''    Ctrl.State = False
'
''---Открываем окно свойств обращенной фигуры
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---Выходим из процедуры обработки
'    Application.EventsEnabled = True
'End Sub
'
'Public Sub MakeVHoseLine()
''Метод обращения во всасывающую рукавную линию
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
'    On Error GoTo Tail
'
''---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
'    Application.EventsEnabled = False
'    ImportVHoseInformation
'    Application.EventsEnabled = True
'
''---Идентифицируем активную фигуру
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---Получаем списки для фигуры
'    '---Запускаем процедуру получения списка Подразделений
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
''    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
''    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал")
''    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
''    HoseDiametersListImport (ShpInd)
''    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
''    HoseResistanceValueImport (ShpInd)
''    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
''    HoseMaxFlowValueImport (ShpInd)
''    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
''    HoseWeightValueImport (ShpInd)
'
''---Устанавливаем значение текущего времени бкз ссылки
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---Открываем окно свойств обращенной фигуры
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---Выходим из процедуры обработки
'    Application.EventsEnabled = True
'End Sub
'
'Public Sub MakeMagHoseLine()
''Метод обращения в Магистральную рукавную линию
'Dim ShpObj As Visio.Shape
'Dim ShpInd As Integer
'
''---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
'    On Error GoTo Tail
'
''---Отключаем обработку событий приложением, обращаем фигуру и вновь включаем обработку событий
'    Application.EventsEnabled = False
'    ImportHoseInformation
'    Application.EventsEnabled = True
'
''---Идентифицируем активную фигуру
'    Set ShpObj = Application.ActiveWindow.Selection(1)
'    ShpInd = ShpObj.ID
'
''---Получаем списки для фигуры
'    '---Запускаем процедуру получения списка Подразделений
'    ShpObj.Cells("Prop.Unit.Format").FormulaU = ListImport("Подразделения", "Подразделение")
'    '---Запускаем процедуру получения СПИСКОВ Материалов рукавов
'    ShpObj.Cells("Prop.HoseMaterial.Format").FormulaU = ListImport("З_Рукава", "Материал рукава")
'    '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
'    HoseDiametersListImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
'    HoseResistanceValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
'    HoseMaxFlowValueImport (ShpInd)
'    '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
'    HoseWeightValueImport (ShpInd)
'
''---Устнавливаем значения для магистральной линии
'    ShpObj.Cells("Prop.HoseDiameter").FormulaU = "INDEX(2,Prop.HoseDiameter.Format)"
'    ShpObj.Cells("Prop.LineType").FormulaU = "INDEX(1,Prop.LineType.Format)"
'
''---Устанавливаем значение текущего времени бкз ссылки
'    ShpObj.Cells("Prop.LineTime").FormulaU = _
'        "DATETIME(" & Str(ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)) & ")"
'
''---Открываем окно свойств обращенной фигуры
'    On Error Resume Next
'    Application.DoCmd (1312)
'
'Exit Sub
'Tail:
'    '---Выходим из процедуры обработки
'    Application.EventsEnabled = True
'End Sub
Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Расчет НРС
    
    'NormalizeNRS
Dim NRSDemon As c_NRSDemon
Dim shp As Visio.Shape

    On Error GoTo ex
    
    If Not Application.ActiveWindow.Selection.Count = 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    
    Set NRSDemon = New c_NRSDemon
'    NRSDemon.roundAccuracy = 4
'    NRSDemon.checkAccuracy = 2
'    NRSDemon.outAccuracy = 2
'    NRSDemon.maxIterations = 100
'    NRSDemon.approvedHout = 3
    'Строим модель
    NRSDemon.BuildNRSModel shp
    'Рассчитываем модель
    NRSDemon.CalculateNRSModel
    
    Set NRSDemon = Nothing
Exit Sub
ex:
    Set NRSDemon = Nothing
End Sub
