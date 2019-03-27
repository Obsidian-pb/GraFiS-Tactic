Attribute VB_Name = "m_ToolBarReport"
Option Explicit
'--------------------------------Модуль для добавления панели инструментов Документа - Отчеты----------------------------


Public Sub AddTBReport()
'Процедура добавления панели управления "Отчеты"-------------------------------
Dim i As Integer

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
'---Проверяем есть ли уже панель управления "Отчеты"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Отчеты" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Отчеты"
        .Visible = True
    End With

End Sub

Public Sub RemoveTBReport()
'Процедура добавления панели управления "Отчеты"-------------------------------
    Application.CommandBars("Отчеты").Delete
End Sub

Public Sub AddButtonMakeReport()
'Процедура добавление новой кнопки на панель управления "Превращения"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Отчеты")
'    DocPath = Application.Documents("Отчеты.vss").Path

'---Добавляем кнопки на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить в рукавную линию"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Сделать отчет"
        .Tag = "Report"
        .TooltipText = "Сфоримровать форму отчета"
        .FaceID = Visio.visIconIXCUSTOM_BOX
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub


Public Sub DeleteButtonReport()
'---Процедура удаления кнопки "Сделать отчет" с панели управления "Отчеты"--------------

'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Отчеты")
'---Удаление кнопки "Сделать отчет" на панели управления "Отчеты"------------------------
    Set Button = Bar.Controls("Сделать отчет")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub

