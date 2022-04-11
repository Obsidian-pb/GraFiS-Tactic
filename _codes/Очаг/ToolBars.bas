Attribute VB_Name = "ToolBars"

Sub AddTBImagination()
'Процедура добавления панели управления "Превращения"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    'Const DocPath = ThisDocument.Path
    'Dim DocPath As String
    
'---Проверяем есть ли уже панель управления "Превращения"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Превращения" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Превращения"
        .Visible = True
    End With

End Sub

Sub RemoveTBImagination()
'Процедура добавления панели управления "Превращения"-------------------------------
    Application.CommandBars("Превращения").Delete
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Превращения"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("Превращения")
    DocPath = ThisDocument.path
    
'---Добавляем кнопки на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить в расчетную зону"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Расчетная зона"
        .tag = "CalcArea"
        .TooltipText = "Обратить в расчетную зону"
'        .Picture = LoadPicture(DocPath & "Bitmaps\Fire1.bmp")
'        .Mask = LoadPicture(DocPath & "Bitmaps\Fire2.bmp")
        .FaceID = 150
        .BeginGroup = True
    End With
'---Кнопка "Обратить в зону горения"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Площадь"
        .tag = "FireAreae"
        .TooltipText = "Обратить в зону горения"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fire1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fire2.bmp")
        .BeginGroup = False
    End With
'---Кнопка "Обратить в зону горения"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Шторм"
        .tag = "FireStorm"
        .TooltipText = "Обратить в огненный шторм"
        .Picture = LoadPicture(DocPath & "Bitmaps\Storm1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Storm2.bmp")
    End With
'---Кнопка "Обратить в Задымление"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Задымление"
        .tag = "Fog"
        .TooltipText = "Обратить в задымленную зону"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fog1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fog2.bmp")
    End With
'---Кнопка "Обратить в зону обрушения"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Обрушение"
        .tag = "Rush"
        .TooltipText = "Обратить в зону обрушения"
        .Picture = LoadPicture(DocPath & "Bitmaps\Rush1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Rush2.bmp")
    End With
    
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Площадь" из панели управления "Превращения"--------------
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error Resume Next

    Set Bar = Application.CommandBars("Превращения")
'---Удаление кнопки "Расчетная зона" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Расчетная зона")
    Button.Delete
'---Удаление кнопки "Площадь" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Площадь")
    Button.Delete
'---Удаление кнопки "Шторм" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Шторм")
    Button.Delete
'---Удаление кнопки "Задымление" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Задымление")
    Button.Delete
'---Удаление кнопки "Обрушение" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Обрушение")
    Button.Delete
    
    
Set Button = Nothing
Set Bar = Nothing

End Sub

'-----------------------------Инструменты работы с кнопками-------------------------------------
Public Sub PS_CheckButtons(ByRef a_MainBtn As Office.CommandBarButton)
'Процедура Оставляет включенной только указанную кнопку (в том случае, если не выбрано ни одной фигуры)
Dim v_Cntrl As CommandBarControl
    
    If Application.ActiveWindow.Selection.Count >= 1 Then Exit Sub
    
    For Each v_Cntrl In Application.CommandBars("Превращения").Controls
        If v_Cntrl.Caption = a_MainBtn.Caption Then
            v_Cntrl.State = Not a_MainBtn.State 'msoButtonDown
        Else
            v_Cntrl.State = False 'msoButtonUp
        End If
    Next v_Cntrl
End Sub
