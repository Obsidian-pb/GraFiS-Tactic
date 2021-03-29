Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTBImagination()
'Процедура добавления панели управления "РТП"-------------------------------
Dim i As Integer
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---Проверяем есть ли уже панель управления "РТП"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "РТП" Then Exit Sub
    Next i

'---Создаем панель управления "РТП"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "РТП"
        .Visible = True
    End With

End Sub

Sub RemoveTBImagination()
'Процедура добавления панели управления "РТП"-------------------------------
    On Error Resume Next
    Application.CommandBars("РТП").Delete
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "РТП"--------------
    
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("РТП")
    
'---Добавляем кнопки на панель управления "РТП"--------------------------------
'---Кнопка "Команда"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Команда"
        .Tag = "Command"
        .TooltipText = "Команда тактической единице"
        .FaceID = 238
    End With
    
    Set btns = New c_Buttons
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons", "Команды"
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Команда" из панели управления "РТП"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("РТП")
'---Удаление кнопки "Команда" из панели управления "РТП"------------------------
    Set Button = Bar.Controls("Команда")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
EX:
'Выходим из процедуры
End Sub
