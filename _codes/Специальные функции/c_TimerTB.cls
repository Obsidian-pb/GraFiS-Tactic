VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_TimerTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Класс для работы панели инструментов "Таймер"

Public TB_Date As Office.CommandBarControl  'Дата
Public TB_Time As Office.CommandBarControl  'Время
Public WithEvents ComButMinuteReduce As Office.CommandBarButton  'Уменьшить на минуту
Attribute ComButMinuteReduce.VB_VarHelpID = -1
Public WithEvents ComButMinuteRize As Office.CommandBarButton  'Увеличить на минуту
Attribute ComButMinuteRize.VB_VarHelpID = -1
Public WithEvents ComButCurrentTime As Office.CommandBarButton  'Текущее время
Attribute ComButCurrentTime.VB_VarHelpID = -1
Public WithEvents ComButSystemTime As Office.CommandBarButton  'Системное время
Attribute ComButSystemTime.VB_VarHelpID = -1
Public WithEvents ComButFireTime As Office.CommandBarButton  'Время начала пожара
Attribute ComButFireTime.VB_VarHelpID = -1
Public WithEvents ComButLastTime As Office.CommandBarButton  'Самое позднее время
Attribute ComButLastTime.VB_VarHelpID = -1
Public WithEvents ComButShapeTime As Office.CommandBarButton  'Время выбранной фигуры
Attribute ComButShapeTime.VB_VarHelpID = -1
Public WithEvents ComButTimeToShape As Office.CommandBarButton  'Установить текущее время
Attribute ComButTimeToShape.VB_VarHelpID = -1

Public CurrentTimerActive As Boolean
'----------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    '---Добавляем тулбар таймера
    AddTBTimerDirect
    '---Ассоциируем кнопки
    Set ComButMinuteReduce = Application.CommandBars("Таймер").Controls("Уменьшить на 1 минуту")
    Set ComButMinuteRize = Application.CommandBars("Таймер").Controls("Увеличить на 1 минуту")
    Set ComButSystemTime = Application.CommandBars("Таймер").Controls("Системное время")
    Set ComButCurrentTime = Application.CommandBars("Таймер").Controls("Текущее время")
    Set ComButFireTime = Application.CommandBars("Таймер").Controls("Время начала пожара")
    Set ComButLastTime = Application.CommandBars("Таймер").Controls("Самое позднее время")
    Set ComButShapeTime = Application.CommandBars("Таймер").Controls("Время выбранной фигуры")
    Set ComButTimeToShape = Application.CommandBars("Таймер").Controls("Установить текущее время")
    '---Ассоциируем поля
    Set TB_Date = Application.CommandBars("Таймер").Controls("Дата")
    Set TB_Time = Application.CommandBars("Таймер").Controls("Время")
    '---Отключаем таймер по-умолчанию
    CurrentTimerActive = False
    '---Включаем общий таймер
    tmrStart
End Sub

Private Sub Class_Terminate()
    '---Деактивируем кнопки
    RemoveTBTimer
End Sub

'----------------------------------------------------------------------------------------
Public Sub AddTBTimerDirect()
'---Добавляем тулбокс таймера
Dim i As Integer

    On Error GoTo ex

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

'---Проверяем есть ли уже панель управления "Таймер"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Таймер" Then Exit Sub
    Next i

'---Создаем панель управления "Цветовые схемы"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True)
    With Bar
        .Name = "Таймер"
        .Visible = True
    End With

    AddTBTimerButtons

Exit Sub
ex:
    'Error

End Sub

Public Sub RemoveTBTimer()
'Процедура добавления панели управления "Таймер"-------------------------------
    On Error GoTo ex
'---Отключаем таймер текущего системного времени, если он был включен
    tmrKill
'---Удаляем панель управления
    Application.CommandBars("Таймер").Delete
ex:
End Sub

Public Sub AddTBTimerButtons()
'---Добавляем кнопки на тулбокс таймера
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, control As CommandBarControl, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("Таймер")
    
    Dim tbc As c_TimerTB
    
'---Добавляем кнопки на панель управления "Таймер"--------------------------------
'---Контрол "Дата"
    Set control = Bar.Controls.Add(Type:=msoControlEdit)
    With control
        .Caption = "Дата"
        .Tag = "CurrentDate"
        .TooltipText = "Текущая дата схемы"
        .Width = 55
'        .OnAction = "OnCurrentDateAction"
    End With
    PS_LoadFromSheetDate control
    Set control = Nothing
'---Кнопка "Увеличить на минуту"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Уменьшить на 1 минуту"
        .Tag = "Decrease1Min"
        .TooltipText = "Уменьшить текущее время схемы на 1 минуту"
        .FaceID = 2944
        .BeginGroup = True
    End With
    Set Button = Nothing
'---Контрол "Дата"
    Set control = Bar.Controls.Add(Type:=msoControlEdit)
    With control
        .Caption = "Время"
        .Tag = "CurrentTime"
        .TooltipText = "Текущее время схемы"
        .Width = 45
    End With
    PS_LoadFromSheetTime control
    Set control = Nothing
'---Кнопка "Уменьшить на минуту"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Увеличить на 1 минуту"
        .Tag = "Rise1min"
        .TooltipText = "Увеличить текущее время схемы на 1 минуту"
        .FaceID = 2945
    End With
    Set Button = Nothing
'---Кнопка "Системное время"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Системное время"
        .Tag = "SystemTime"
        .TooltipText = "Переключиться на системеное время (разово)"
        .FaceID = 2743
        .BeginGroup = True
'        .OnAction = "OnSystemTimeAction"
    End With
    Set Button = Nothing
'---Кнопка "Текущее время"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Текущее время"
        .Tag = "CurrentTimeContinous"
        .TooltipText = "Переключиться на текущее системеное время (таймер с интервалом 10с.)"
        .FaceID = 2146
    End With
    Set Button = Nothing
'---Кнопка "Время начала пожара"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Время начала пожара"
        .Tag = "FireTime"
        .TooltipText = "Переключиться на время начала пожара"
        .FaceID = 1885
    End With
    Set Button = Nothing
'---Кнопка "Самое позднее время"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Самое позднее время"
        .Tag = "LastTime"
        .TooltipText = "Переключиться на время самой поздней фигуры"
        .FaceID = 157
    End With
    Set Button = Nothing
'---Кнопка "Время выбранной фигуры"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Время выбранной фигуры"
        .Tag = "ShapeTime"
        .TooltipText = "Переключиться на время выбранной фигуры"
        .FaceID = 110
    End With
    Set Button = Nothing
'---Кнопка "Установить текущее время"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Установить текущее время"
        .Tag = "TimeToShape"
        .TooltipText = "Установить для выбранной фигуры текущее время схемы"
        .FaceID = 111
        .BeginGroup = True
    End With
    Set Button = Nothing


Set Bar = Nothing
Exit Sub
ex:
'    Debug.Print Err.Description
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddTBTimerButtons"
End Sub

Public Sub RemooveTBTimer()
'Процедура добавления панели управления "Таймер"
    Application.CommandBars("Таймер").Delete
End Sub




'-------------------------Обработчики событий кнопок----------------------------------------------------
Private Sub ComButMinuteReduce_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnDelMinuteAction
End Sub
Private Sub ComButMinuteRize_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnAddMinuteAction
End Sub
Private Sub ComButSystemTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnSystemTimeAction
End Sub
Private Sub ComButCurrentTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnCurrentTimeComtinousAction
End Sub
Private Sub ComButFireTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnFireTimeAction
End Sub
Private Sub ComButLastTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnLastTimeAction
End Sub
Private Sub ComButShapeTime_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnShapeTimeAction
End Sub
Private Sub ComButTimeToShape_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    OnShapeTimeToShapeAction
End Sub






'-------------------------Проки и функции по реагированию на действия с контролами----------------------
Public Sub OnCurrentDateAction()
'Прока запускаемая при ввдении данных в поле CurrentDate
'---Отключаем таймер (если был включен)
'    tmrKill
    On Error GoTo ex
'---Проверяем корректность введенного значения
    If IsDate(TB_Date.Text) Then
    '---Если корректно - обновляем значение ячейки TheDoc!User.CurrentTime
        PS_UpdateDateTime TB_Date, TB_Time
    Else
    '---Если некорректно - возвращаем то, что храниться в ячейке TheDoc!User.CurrentTime
        PS_LoadFromSheetDate TB_Date
    End If
    
ex:
End Sub

Public Sub OnCurrentTimeAction()
'Прока запускаемая при ввдении данных в поле CurrentTime
    
'---Отключаем таймер (если был включен)
'    tmrKill
    On Error GoTo ex
'---Проверяем корректность введенного значения
    If IsDate(TB_Time.Text) Then
    '---Если корректно - обновляем значение ячейки TheDoc!User.CurrentTime
        PS_UpdateDateTime TB_Date, TB_Time
    Else
    '---Если некорректно - возвращаем то, что храниться в ячейке TheDoc!User.CurrentTime
        PS_LoadFromSheetTime TB_Time
    End If
    
ex:
End Sub

Private Sub OnDelMinuteAction()
'Прока запускается при нажатии на кнопку "Уменьшить на 1 минуту"
Dim tmpDate As Date
    
    On Error GoTo ex
    
'---Отключаем таймер (если был включен)
    CurrentTimerActive = False

'---Получаем текущее значение
    tmpDate = TimeValue(TB_Time.Text)
    
'---Уменьшаем его на 1 минуту
    tmpDate = DateAdd("n", -1, tmpDate)
    
'---Устанавливаем новое значение для контрола
    TB_Time.Text = TimeValue(tmpDate)
    
'---Если новое значение = 59минут то уменьшаем дату на 1 день
    If Hour(tmpDate) = 23 And Minute(tmpDate) = 59 Then
        PS_DateChange TB_Date, -1
    End If
    
'---Передаем данные в контрол
    PS_UpdateDateTime TB_Date, TB_Time

ex:
End Sub
Private Sub OnAddMinuteAction()
'Прока запускается при нажатии на кнопку "Увеличить на 1 минуту"
Dim tmpDate As Date

    On Error GoTo ex

'---Отключаем таймер (если был включен)
    CurrentTimerActive = False

'---Получаем текущее значение
    tmpDate = TimeValue(TB_Time.Text)
    
'---Уменьшаем его на 1 минуту
    tmpDate = DateAdd("n", 1, tmpDate)
    
'---Устанавливаем новое значение для контрола
    TB_Time.Text = TimeValue(tmpDate)
    
'---Если новое значение = 59минут то уменьшаем дату на 1 день
    If Hour(tmpDate) = 0 And Minute(tmpDate) = 0 Then
        PS_DateChange TB_Date, 1
    End If
    
'---Передаем данные в контрол
    PS_UpdateDateTime TB_Date, TB_Time

ex:
End Sub

Private Sub OnSystemTimeAction()
'Прока запускаетсяпри нажатии на кнопку "Переключиться на системное время"

    On Error GoTo ex

'---Отключаем таймер (если был включен)
    CurrentTimerActive = False

'---Устанавливаем значения даты и времени для всех полей
    TB_Date.Text = DateValue(Now())
    TB_Time.Text = TimeValue(Now())
    
'---Передаем данные в контрол
    PS_UpdateDateTime TB_Date, TB_Time
    
ex:
End Sub

Private Sub OnFireTimeAction()
'Прока запускается при нажатии на кнопку "FireTime"

    On Error GoTo ex

'---Отключаем таймер (если был включен)
    CurrentTimerActive = False
    
'---Устнаваливаем ссылку для ячейки "TheDoc!User.CurrentTime"
    If GetFireTime Then
        Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
            "User.FireTime"
    Else
        MsgBox "На текущем рабочем листе нет фигур Очага горения"
        Exit Sub
    End If

    
'---Устанавливаем значения даты и времени для всех полей
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnCurrentTimeComtinousAction()
'Прока запускается при нажатии на кнопку "CurrentTime"

    On Error GoTo ex

'---Показываем текуще значение времени
    OnSystemTimeAction

'---Включаем таймер (если был выключен)
'    tmrStart
    CurrentTimerActive = True
    
ex:
End Sub

Private Sub OnLastTimeAction()
'Прока запускается при нажатии на кнопку "FireTime"

    On Error GoTo ex

'---Отключаем таймер (если был включен)
    CurrentTimerActive = False
    
'---Устнаваливаем ссылку для ячейки "TheDoc!User.CurrentTime"
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        GetLastTime(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.FireTime").Result(visDate))
    
'---Устанавливаем значения даты и времени для всех полей
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnShapeTimeAction()
'Прока запускается при нажатии на кнопку "ShapeTime" - получить время выбранной фигуры

    On Error GoTo ex
    
'---Если ни одна фигура не выделена - выходим
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
'---Отключаем таймер (если был включен)
    CurrentTimerActive = False
    
'---Устнаваливаем значение для ячейки "TheDoc!User.CurrentTime"
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        GetCurrentShapeTime(Application.ActiveWindow.Selection(1))
    
'---Устанавливаем значения даты и времени для всех полей
    PS_LoadFromSheetDate TB_Date
    PS_LoadFromSheetTime TB_Time
    
ex:
End Sub

Private Sub OnShapeTimeToShapeAction()
'Прока запускается при нажатии на кнопку "TimeToShape" - Установить для выбранной фигуры (фигур) текущее время схемы
Dim shp As Visio.Shape
Dim curTime As Double

    On Error GoTo ex
    
'---Если ни одна фигура не выделена - выходим
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
'---Получаем текущее время схемы
    curTime = Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Result(visDate)
    
'---Перебираем все фигуры в выделении
    For Each shp In Application.ActiveWindow.Selection
        SetCurrentShapeTime shp, curTime
    Next shp
    
ex:
End Sub


'---------------Служебные проки------------------------------------------------------------------
Private Sub PS_DateChange(ByRef controlDate As CommandBarControl, ByVal changeValue As Double)
'Прока уменьшает значение даты в контроле CurrentDate на 1 день
Dim tmpDate As Date

    On Error GoTo ex

'---Получаем текущее значение
    tmpDate = DateValue(controlDate.Text)
    
'---Уменьшаем его на 1 день
    tmpDate = DateAdd("d", changeValue, tmpDate)
    
'---Устанавливаем новое значение для контрола
    controlDate.Text = DateValue(tmpDate)
    
ex:
End Sub

Public Sub FillFullData()
'Если в полях времени и даты пусто - устанавливаем

    On Error GoTo ex

'    Debug.Print TB_Date.Text
    If TB_Date.Text = "" Then
        PS_LoadFromSheetDate TB_Date
    End If
    If TB_Time.Text = "" Then
        PS_LoadFromSheetTime TB_Time
    End If
    
ex:
End Sub

Private Function GetLastTime(ByVal fireStartTime As Double) As String
'Функция возвращает значение последнего установленного времени
Dim shp As Visio.Shape
Dim curVal As Double
Dim maxVal As Double
    
    On Error GoTo ex
    
    maxVal = fireStartTime
    '---Перебираем все фигуры на листе и если значение ячайки со временм яляется максимальным возвращаем его
    For Each shp In Application.ActivePage.Shapes
        curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
                 GetShapeTime(shp, "Prop.LineTime") + _
                 GetShapeTime(shp, "Prop.SetTime") + _
                 GetShapeTime(shp, "Prop.SquareTime") + _
                 GetShapeTime(shp, "Prop.StabCreationTime") + _
                 GetShapeTime(shp, "Prop.UTPCreationTime") + _
                 GetShapeTime(shp, "Prop.FormingTime") + _
                 GetShapeTime(shp, "Prop.FindTime") + _
                 GetShapeTime(shp, "Prop.RushTime")
        If curVal > maxVal And curVal < maxVal + 100 Then maxVal = curVal 'Здесь +100 -> учет перебора для фигур у которых несколько ячеек с именами времен, например звенья ГДЗС
    Next shp
    
    GetLastTime = "DATETIME(" & str(maxVal) & ")"
Exit Function
ex:
    GetLastTime = "User.FireTime"
End Function
Private Function GetShapeTime(ByRef shp As Visio.Shape, ByVal cellName) As Double
'Функция возвращает значение для текущей ячейки, если такой ячейки нет - возвращается 0
On Error GoTo ex
    GetShapeTime = shp.Cells(cellName).Result(visDate)
Exit Function
ex:
    GetShapeTime = 0
End Function

Private Function GetCurrentShapeTime(ByRef shp As Visio.Shape) As String
'Функция возвращает значение времени для выбранной фигуры
Dim curVal As Double
    
    On Error GoTo ex

'---Получаем значение любой ячейки содержащей значения времени прибытия (установки и т.д.)
    curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
         GetShapeTime(shp, "Prop.LineTime") + _
         GetShapeTime(shp, "Prop.SetTime") + _
         GetShapeTime(shp, "Prop.SquareTime") + _
         GetShapeTime(shp, "Prop.StabCreationTime") + _
         GetShapeTime(shp, "Prop.UTPCreationTime") + _
         GetShapeTime(shp, "Prop.FormingTime") + _
         GetShapeTime(shp, "Prop.FindTime") + _
         GetShapeTime(shp, "Prop.RushTime") + _
         GetShapeTime(shp, "Prop.FireTime")
    
    curVal = GetGDZSChainCorrectTime(shp, curVal)
    
    If curVal = 0 Then
        MsgBox "Данная фигура не обладает свойствами времени! Значение текущего времени будет поставлено равным времени начала пожара.", vbInformation, "Предупреждение"
        GetCurrentShapeTime = "User.FireTime"
    Else
        GetCurrentShapeTime = "DATETIME(" & str(curVal) & ")"
    End If

Exit Function
ex:
    GetCurrentShapeTime = "User.FireTime"
End Function

Private Sub SetCurrentShapeTime(ByRef shp As Visio.Shape, ByRef dateVal As Double)
'Процедура устанавливает значение времени для выбранной фигуры
    On Error Resume Next
    With shp
        .Cells("Prop.ArrivalTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.LineTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.SetTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.SquareTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.StabCreationTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.UTPCreationTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FormingTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FindTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.RushTime").Formula = "DATETIME(" & dateVal & ")"
        .Cells("Prop.FireTime").Formula = "DATETIME(" & dateVal & ")"
    End With
 
End Sub

Private Function GetGDZSChainCorrectTime(ByRef shp As Visio.Shape, ByVal testVal As Double) As Double
'Функция возвращае ткорректное значение для фигур звеньев ГДЗС
On Error GoTo ex
    If shp.CellExists("User.IndexPers", 0) Then
        If shp.Cells("User.IndexPers").Result(visNumber) = 46 Then
            GetGDZSChainCorrectTime = shp.Cells("Prop.FormingTime").Result(visDate)
            Exit Function
        End If
    End If
ex:
    GetGDZSChainCorrectTime = testVal
End Function

Private Function GetFireTime() As Boolean
'Функция устанавливает значение времени начала пожара для ячейки User.FireTime и вщзвращает истину если фигура есть
'Ложь - если нет
Dim shp As Visio.Shape
Dim timeVal As Double
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers").Result(visNumber) = 70 Then
                timeVal = GetShapeTime(shp, "Prop.FireTime")
                Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.FireTime").Formula = _
                    "DATETIME(" & str(timeVal) & ")"
                
                GetFireTime = True
                Exit Function
            End If
        End If
    Next shp

GetFireTime = False
Exit Function
ex:
    GetFireTime = False
End Function



'---------------Проки обмена данными------------------------------------------------------------
Public Sub PS_UpdateDateTime(ByRef controlDate As CommandBarControl, ByRef controlTime As CommandBarControl)
'Обновляем значение поля TheDoc!User.CurrentTime в соответсвии с переданным значением
Dim tmpDate As Date
    On Error GoTo ex
    tmpDate = DateValue(controlDate.Text) + TimeValue(controlTime.Text)
    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
        "DATETIME(" & str(CDbl(tmpDate)) & ")"

ex:
End Sub

Public Sub PS_LoadFromSheetDate(ByRef control As CommandBarControl)
'Устанавливаем полю контрола ДАТЫ значение которое было в ячейке TheDoc!User.CurrentTime (отменяем изменения)
    On Error Resume Next
    control.Text = _
            Left(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").ResultStr(0), 10)
End Sub

Public Sub PS_LoadFromSheetTime(ByRef control As CommandBarControl)
'Устанавливаем полю контрола ВРЕМЕНИ значение которое было в ячейке TheDoc!User.CurrentTime (отменяем изменения)
    On Error Resume Next
    control.Text = _
            Mid(Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").ResultStr(0), 12)
End Sub





'Private Sub VisioApp_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
''Debug.Print GetKeyState(16)
''GetKeyState (VKey)
'End Sub
'
''-------------Учет времени выбранной фигуры----------------------
'Private Sub VisioApp_SelectionChanged(ByVal Window As IVWindow)
''Реагируем на измненеия Selection
'Dim val As Integer
'On Error Resume Next
'    val = GetKeyState(16)
''Debug.Print val
''    If Window.Selection.Count = 0 Then Exit Sub
'    If val <> 0 Then OnSelectionChangeAction Window.Selection(1)
'End Sub
'
'Private Sub OnSelectionChangeAction(ByRef shp As Visio.Shape)
''Прока запускается при нажатии на кнопку "FireTime"
'Dim curVal As Double
'
'    On Error GoTo EX
'
''---Отключаем таймер (если был включен)
'    CurrentTimerActive = False
'
''---Получаем время данной фигуры
'    curVal = GetShapeTime(shp, "Prop.ArrivalTime") + _
'         GetShapeTime(shp, "Prop.LineTime") + _
'         GetShapeTime(shp, "Prop.SetTime") + _
'         GetShapeTime(shp, "Prop.SquareTime") + _
'         GetShapeTime(shp, "Prop.StabCreationTime") + _
'         GetShapeTime(shp, "Prop.UTPCreationTime") + _
'         GetShapeTime(shp, "Prop.FormingTime") + _
'         GetShapeTime(shp, "Prop.FindTime") + _
'         GetShapeTime(shp, "Prop.RushTime")
'
''---Устнаваливаем ссылку для ячейки "TheDoc!User.CurrentTime"
'    Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime").Formula = _
'        "DATETIME(" & Str(curVal) & ")"
'
''---Устанавливаем значения даты и времени для всех полей
'    PS_LoadFromSheetDate TB_Date
'    PS_LoadFromSheetTime TB_Time
'
'Exit Sub
'EX:
'    Debug.Print Err.Description
'End Sub

