Attribute VB_Name = "m_Toolbars"
Sub AddTB_Constructions()
'Процедура добавления панели управления "Конструкции"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---Проверяем есть ли уже панель управления "Конструкции"------------------------------
    For i = 1 To Application.CommandBars.count
        If Application.CommandBars(i).Name = "Конструкции" Then Exit Sub
    Next i

'---Создаем панель управления "Конструкции"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Конструкции"
        .Visible = True
    End With
'    AddButtons
End Sub

Sub RemoveTB_Constructions()
'Процедура добавления панели управления "Конструкции"-------------------------------
    On Error GoTo EX
    Application.CommandBars("Конструкции").Delete

EX:
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Конструкции"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Конструкции")
    DocPath = Application.Documents("Конструкции.vss").path
    
'---Добавляем кнопки на панель управления "Конструкции"--------------------------------
'---Кнопка "Маска"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .caption = "Маска"
        .Tag = "WallsMask"
        .TooltipText = "Наложить маску стен"
        .Picture = LoadPicture(DocPath & "Bitmaps\WallMask1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallMask2.bmp")
    End With
'---Кнопка "Рисование стен"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .caption = "Рисование стен"
        .Tag = "WallDrawer"
        .TooltipText = "Рисование стен"
        .Picture = LoadPicture(DocPath & "Bitmaps\WallDrawTool1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallDrawTool2.bmp")
    End With

    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "AddButtons"
End Sub
