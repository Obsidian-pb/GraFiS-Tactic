Attribute VB_Name = "ToolBars"



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
'---Кнопка "Экспорт в JPG"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Мастер проверок"
        .Tag = "show_m_chek_form"
        .TooltipText = "Проверить правильность схемы"
        .FaceID = 172
        .BeginGroup = True
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
'---Процедура удаления кнопки "Мастер проверок" из панели управления "Спецфункции"--------------
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Спецфункции")
'---Удаление кнопки "Рукав" из панели управления "Мастер проверок"------------------------
    Set Button = Bar.Controls("Мастер проверок")
    Button.Delete
    
    
Set Button = Nothing
Set Bar = Nothing

End Sub
