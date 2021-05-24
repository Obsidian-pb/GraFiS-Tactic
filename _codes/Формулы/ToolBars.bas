Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTB_f()
'Процедура добавления панели управления "Формулы"-------------------------------
Dim i As Integer
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---Проверяем есть ли уже панель управления "Формулы"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).name = "Формулы" Then Exit Sub
    Next i

'---Создаем панель управления "Формулы"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .name = "Формулы"
        .visible = True
    End With

End Sub

Sub RemoveTB_f()
'Процедура добавления панели управления "Формулы"-------------------------------
    On Error Resume Next
    Application.CommandBars("Формулы").Delete
End Sub

Sub AddButtons_f()
'Процедура добавление новой кнопки на панель управления "Формулы"--------------
    
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("Формулы")
    
'---Добавляем кнопки на панель управления "Формулы"--------------------------------
'---Кнопка "Обновить"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Обновить"
        .Tag = "Refresh all formulas"
        .TooltipText = "Обновить все формулы на листе"
        .FaceID = 37
    End With
'---Кнопка "Показать все"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Показать все"
        .Tag = "Show all formulas"
        .TooltipText = "Показать все вычисления в одном окне"
        .FaceID = 139
    End With
    
    Set btns = New c_Buttons
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
ex:
    Set Button = Nothing
    Set Bar = Nothing
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.name
    SaveLog Err, "AddButtons_f", "Кнопки на панели Формулы"
End Sub


Sub DeleteButtons_f()
'---Процедура удаления кнопок из панели управления "Формулы"--------------
    On Error GoTo ex
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("Формулы")
'---Удаление кнопки "Обновить" из панели управления "Формулы"------------------------
    Set Button = Bar.Controls("Обновить")
    Button.Delete
'---Удаление кнопки "Обновить" из панели управления "Формулы"------------------------
    Set Button = Bar.Controls("Показать все")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
ex:
'Выходим из процедуры
    Set Button = Nothing
    Set Bar = Nothing
    SaveLog Err, "AddButtons_f", "Кнопки на панели Формулы"
End Sub
