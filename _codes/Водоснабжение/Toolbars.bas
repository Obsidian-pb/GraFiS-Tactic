Attribute VB_Name = "Toolbars"
Sub AddTBImagination()
'Процедура добавления панели управления "Превращения"-------------------------------
Dim i As Integer

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---Проверяем есть ли уже панель управления "Превращения"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Превращения" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
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

    Set Bar = Application.CommandBars("Превращения")
    DocPath = Application.Documents("Водоснабжение.vss").path
    
'---Добавляем кнопки на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить в открытый водоисточник"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Естественный водоисточник"
        .Tag = "NaturalWater"
        .TooltipText = "Обратить в естественный водоисточник"
        .Picture = LoadPicture(DocPath & "Bitmaps\Lake1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Lake2.bmp")
    End With
    
    Set Button = Nothing

Set Bar = Nothing
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Площадь" из панели управления "Превращения"--------------

'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
'---Удаление кнопки "Рукав" из панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Естественный водоисточник")
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
