Attribute VB_Name = "m_TimerWork"
Option Explicit

'Public elpTm As Long
'Public tmrID As Long
'Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

#If VBA7 Then
    Public elpTm As LongPtr
    Public tmrID As LongPtr
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIDEvent As LongPtr, _
                    ByVal uElapse As LongPtr, _
                    ByVal lpTimerfunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
                    ByVal hwnd As LongPtr, _
                    ByVal nIDEvent As LongPtr) As LongPtr
#Else
    Public elpTm As Long
    Public tmrID As Long
    Public Declare Function SetTimer Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal nIDEvent As Long, _
                    ByVal uElapse As Long, _
                    ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" ( _
                    ByVal hwnd As Long, _
                    ByVal nIDEvent As Long) As Long
#End If

Private timerTB As c_TimerTB
'--------------------------------------------------------------------------------------------------


Public Sub AddTimer()
    On Error Resume Next
    Set timerTB = New c_TimerTB
    TimerButtonChamgeState msoButtonDown
End Sub

Public Sub DelTBTimer()
    On Error Resume Next
    Set timerTB = Nothing
    TimerButtonChamgeState msoButtonUp
End Sub

Public Sub TimerButtonChamgeState(ByVal state As MsoButtonState)
'Изменяем состояние кнопки "Таймер"
    On Error Resume Next
    Application.CommandBars("Спецфункции").Controls("Таймер").state = state
End Sub

Public Sub AddTBTimer(ShpObj As Visio.Shape)
'---Добавляем тулбокс таймера
Dim i As Integer

    On Error GoTo ex

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

'---Проверяем есть ли уже панель управления "Таймер"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Таймер" Then 'Exit Sub
            DelTBTimer
            ShpObj.Delete
            Exit Sub
        End If
    Next i

'---Создаем панель управления "Цветовые схемы"--------------------------------------------
    AddTimer

    ShpObj.Delete

Exit Sub
ex:
    'Error
    ShpObj.Delete

End Sub




'--------------Проки таймера---------------------------------------------------------------------
Public Sub tmrStart()
'Запускаем таймер для учета текщего времени
    tmrID = SetTimer(&H0, &H0, 1000, AddressOf tmrPrc)  '1000 - 1 сек
End Sub
'Этопришлось вынести в отдельную процедуру,
'т.к. таймер не желал "убиваться" в таймерной процедуре tmrPrc
Public Sub tmrKill()
    KillTimer &H0, tmrID
End Sub

Public Sub tmrPrc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
'Это обработчик таймера - "таймерная" процедура../
Dim controlDate As CommandBarControl
Dim controlTime As CommandBarControl
Dim timeCell As Visio.Cell
Dim curDateTime As Date

    On Error GoTo ex
'---Определяем объект поля CurrentDate
    Set controlDate = Application.CommandBars("Таймер").Controls("Дата")
'---Определяем объект поля CurrentTime
    Set controlTime = Application.CommandBars("Таймер").Controls("Время")
'---Определяем ячейку содержащую текущее время и само время
    Set timeCell = Application.ActiveDocument.DocumentSheet.Cells("TheDoc!User.CurrentTime")
    curDateTime = timeCell.Result(visDate)
    
'    Debug.Print Now() & " - " & controlDate.Text & ", " & controlTime.Text
'---В случае, если поля панели инструментов "Таймер" обнулились - обновляем панель
    If controlDate.Text = controlTime.Text Then
'        Debug.Print "обновляем"
        DelTBTimer
        AddTimer
    End If
    
'---Проверяем, не менялось ли значение TB_Date и TB_Time
    If Not controlTime.Text = TimeValue(curDateTime) Then
        timerTB.OnCurrentTimeAction
    End If
    If Not controlDate.Text = DateValue(curDateTime) Then
        timerTB.OnCurrentDateAction
    End If
    
'---В случае, если активен таймер текущего времени - обрабатываем 10-ти секундный таймер
    If timerTB.CurrentTimerActive Then
        If DateDiff("s", curDateTime, Now()) >= 10 Then
        '---Устанавливаем значения даты и времени для всех полей
            controlDate.Text = DateValue(Now())
            controlTime.Text = TimeValue(Now())
        '---Передаем данные в контрол
            timerTB.PS_UpdateDateTime controlDate, controlTime
        End If
    End If

'---Если поля обнулились - заполняем их
    timerTB.FillFullData

ex:
End Sub


