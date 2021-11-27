Attribute VB_Name = "Tollbars"
Option Explicit

Private btns As c_Buttons


Sub AddTB()
'Процедура добавления панели управления "Экспорт"-------------------------------
Dim i As Integer
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---Проверяем есть ли уже панель управления "Экспорт"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Экспорт" Then Exit Sub
    Next i

'---Создаем панель управления "Экспорт"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Экспорт"
        .Visible = True
    End With
'---Добавляем кнопки
    AddButtons
    
End Sub

Sub RemoveTB()
'Процедура добавления панели управления "Экспорт"-------------------------------
    On Error Resume Next
    DeleteButtons
    Application.CommandBars("Экспорт").Delete
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Экспорт"--------------
    
'---Объявляем переменные и постоянные--------------------------------------------------
Dim Bar As CommandBar
Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("Экспорт")
    
'---Добавляем кнопки на панель управления "Экспорт"--------------------------------
'---Кнопка "Донесение"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Донесение"
        .Tag = "Report"
        .TooltipText = "Сформировать донесение о пожаре"
        .FaceID = 626
    End With
'---Кнопка "КБД"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "КБД"
        .Tag = "KBD"
        .TooltipText = "Сформировать карточку боевых действий"
        .FaceID = 626
    End With
    
    Set btns = New c_Buttons
    
    Set Bar = Nothing

Exit Sub
ex:
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons", "Экспорт"
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Команда" из панели управления "Экспорт"--------------
    On Error GoTo ex
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar

    Set Bar = Application.CommandBars("Экспорт")
'---Удаление кнопки "Команда" из панели управления "Экспорт"------------------------
    Bar.Controls("Донесение").Delete

Set btns = Nothing

Set Bar = Nothing

Exit Sub
ex:
'Выходим из процедуры
End Sub

