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
        .Caption = "Экспорт в JPG"
        .Tag = "Export_JPG"
        .TooltipText = "Экспортировать все листы в JPG"
        .Picture = LoadPicture(DocPath & "Bitmaps\ExportJPG1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\ExportJPG2.bmp")
        .BeginGroup = True
    End With
'---Кнопка "Аспект"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Аспект"
        .Tag = "Aspect"
        .TooltipText = "Изменить аспект"
        .Picture = LoadPicture(DocPath & "Bitmaps\Aspect1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Aspect2.bmp")
    End With
'---Кнопка "Исправить расположение"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Исправить расположение"
        .Tag = "Fix"
        .TooltipText = "Исправить расположение фигур на листе"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fix1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fix2.bmp")
    End With
'---Кнопка "Количество фигур"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Количество фигур"
        .Tag = "Count"
        .TooltipText = "Показать количество фигур в выборке"
        .Picture = LoadPicture(DocPath & "Bitmaps\Count1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Count2.bmp")
    End With
'---Кнопка "Панель таймера"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Таймер"
        .Tag = "Timer"
        .TooltipText = "Показать панель инструментов 'Таймер'"
        .FaceID = 2146
        .BeginGroup = True
    End With
    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Мастер проверок" из панели управления "Спецфункции"--------------

    On Error GoTo EX

'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Спецфункции")
'---Удаление кнопки "Рукав" из панели управления "Экспорт в JPG"------------------------
    Set Button = Bar.Controls("Экспорт в JPG")
    Button.Delete
'---Удаление кнопки "Рукав" из панели управления "Аспект"------------------------
    Set Button = Bar.Controls("Аспект")
    Button.Delete
'---Удаление кнопки "Рукав" из панели управления "Исправить расположение"------------------------
    Set Button = Bar.Controls("Исправить расположение")
    Button.Delete
'---Удаление кнопки "Рукав" из панели управления "Количество фигур"------------------------
    Set Button = Bar.Controls("Количество фигур")
    Button.Delete
'---Удаление кнопки "Рукав" из панели управления "Таймер"------------------------
    Set Button = Bar.Controls("Таймер")
    Button.Delete
    
    
Set Button = Nothing
Set Bar = Nothing
Exit Sub

EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "DeleteButtons"
End Sub


