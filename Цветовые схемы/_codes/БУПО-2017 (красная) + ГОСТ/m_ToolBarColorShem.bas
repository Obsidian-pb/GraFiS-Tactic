Attribute VB_Name = "m_ToolBarColorShem"
Option Explicit
'--------------------------------Модуль для добавления панели инструментов Документа - Цветовая схема----------------------------


Public Sub AddTBColorShem()
'Процедура добавления панели управления "Цветовые схемы"-------------------------------
Dim i As Integer

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
'---Проверяем есть ли уже панель управления "Цветовые схемы"------------------------------
'    For i = 1 To Application.CommandBars.Count
'        If Application.CommandBars(i).Name = "Цветовые схемы" Then Exit Sub
'    Next i

'---Создаем панель управления "Цветовые схемы"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Цветовые схемы"
        .Visible = True
    End With

End Sub

Public Sub RemoveTBColorShem()
'Процедура удаления панели управления "Цветовые схемы"-------------------------------
    Application.CommandBars("Цветовые схемы").Delete
End Sub

Public Sub AddButtonRefresh()
'Процедура добавление новой кнопки на панель управления "Цветовые схемы"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Цветовые схемы")

'---Добавляем кнопки на панель управления "Цветовые схемы"--------------------------------
'---Кнопка "Обновить цветовую схему"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Обновить"
        .Tag = "Refresh"
        .TooltipText = "Обновить цветовую схему"
'        .FaceID = Visio.visIconIXCUSTOM_BOX
        .FaceID = 625
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub


Public Sub DeleteButtonRefresh()
'---Процедура удаления кнопки "Обновить цветовую схему" с панели управления "Цветовые схемы"--------------

'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Цветовые схемы")
'---Удаление кнопки "Сделать отчет" на панели управления "Отчеты"------------------------
    Set Button = Bar.Controls("Обновить")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub


'-----------------------------Пробные кнопки------------------------------------------

Public Sub AddButtonColorDrop()
'Процедура добавление новой кнопки на панель управления "Цветовые схемы"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Цветовые схемы")

'---Добавляем кнопки на панель управления "Цветовые схемы"--------------------------------
'---Кнопка "Обновить цветовую схему"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=visCtrlTypeSPLITBUTTON, ID:=1692)
'    Set Button = Application.CommandBars("Цветовые схемы").Controls("Цвет")
    With Button
        .Caption = "Цвет"
        .Tag = "ColorDrop"
        .TooltipText = "Цвета"
'        .FaceID = Visio.visIconIXCUSTOM_BOX
        .FaceID = 3
    End With

    Set Button = Nothing

Set Bar = Nothing
End Sub


Public Sub DeleteButtonColor()
'---Процедура удаления кнопки "Обновить цветовую схему" с панели управления "Цветовые схемы"--------------

'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Цветовые схемы")
'---Удаление кнопки "Сделать отчет" на панели управления "Отчеты"------------------------
    Set Button = Bar.Controls("Цвет")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub


Public Sub Prov()
    AddButtonColorDrop 'Пробная


End Sub

Public Sub DeProv()
    DeleteButtonColor 'Пробная


End Sub




