Attribute VB_Name = "ToolBars"

Sub AddTB_SpecFunc()
'Процедура добавления панели управления "Спецфункции"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---Проверяем есть ли уже панель управления "Превращения"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Спецфункции" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Спецфункции"
        .Visible = True
    End With
'    AddButtons
End Sub

Sub RemoveTB_SpecFunc()
'Процедура добавления панели управления "Спецфункции"-------------------------------
    On Error GoTo EX
    Application.CommandBars("Спецфункции").Delete

EX:
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Спецфункции"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Спецфункции")
    DocPath = ThisDocument.path
    
'---Добавляем кнопки на панель управления "Спецфункции"--------------------------------
'---Кнопка "Экспорт в JPG"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Мастер проверок"
        .Tag = "show_m_chek_form"
        .TooltipText = "Проверить правильность схемы"
        .Picture = LoadPicture(DocPath & "Bitmaps\MasterCheck.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MasterCheck2.bmp")
    End With

    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "AddButtons"
End Sub

