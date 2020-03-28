Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons

Sub AddTB_SpecFunc()
'Процедура добавления панели управления "Спецфункции"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---Проверяем есть ли уже панель управления "Спецфункции"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Спецфункции" Then Exit Sub
    Next i

'---Создаем панель управления "Спецфункции"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Спецфункции"
        .Visible = True
    End With

End Sub

Sub RemoveTB_SpecFunc()
'Процедура добавления панели управления "Спецфункции"-------------------------------
    On Error GoTo EX
    Application.CommandBars("Спецфункции").Delete
    Set btns = Nothing
EX:
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Спецфункции"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Спецфункции")
    
'---Добавляем кнопки на панель управления "Спецфункции"--------------------------------
'---Кнопка "Мастер проверок"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Мастер проверок"
        .Tag = "show_m_chek_form"
        .TooltipText = "Мастер проверок"
        .FaceID = 1820
        .beginGroup = True
    End With
'---Кнопка "Тактические данные"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Тактические данные"
        .Tag = "show_m_tactic_form"
        .TooltipText = "Тактические данные"
        .FaceID = 1090
        .beginGroup = False
    End With
'---Кнопка "Настройки"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Настройки"
        .Tag = "calculationSettings"
        .TooltipText = "Настройки"
        .FaceID = 642
        .beginGroup = False
    End With
    
    
    Set Button = Nothing
    
'---Активируем класс отслеживающий кнопку
    Set btns = New c_Buttons
    
Set Bar = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "AddButtons"
End Sub

Sub DeleteButtons()
'---Процедура удаления кнопок с панели управления "Спецфункции"--------------
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("Спецфункции")
'---Удаление кнопки "Мастер проверок" с панели управления "Спецфункции"------------------------
    Set Button = Bar.Controls("Мастер проверок")
    Button.Delete
'---Удаление кнопки "Тактические данные" с панели управления "Спецфункции"------------------------
    Set Button = Bar.Controls("Тактические данные")
    Button.Delete
'---Удаление кнопки "Настройки" с панели управления "Спецфункции"------------------------
    Set Button = Bar.Controls("Настройки")
    Button.Delete
    
'---Активируем класс отслеживающий кнопку
    Set btns = Nothing
    
Set Button = Nothing
Set Bar = Nothing
End Sub
