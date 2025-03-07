Attribute VB_Name = "ToolBars"

Sub AddTBImagination()
'Процедура добавления панели управления "Моделирование"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    'Const DocPath = ThisDocument.Path
    'Dim DocPath As String
    
'---Проверяем есть ли уже панель управления "Превращения"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Моделирование" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Моделирование"
        .Visible = True
    End With

End Sub

Sub RemoveTBImagination()
'Процедура добавления панели управления "Моделирование"-------------------------------
    Application.CommandBars("Моделирование").Delete
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Моделирование"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("Моделирование")
    DocPath = ThisDocument.path
    
'---Добавляем кнопки на панель управления "Моделирование"--------------------------------
'---Кнопка "Обратить в расчетную зону"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Расчетная зона"
        .tag = "CalcAreaMod"
        .TooltipText = "Обратить в расчетную зону"
        .FaceID = 150
        .BeginGroup = True
    End With
'---Кнопка "Моделировать"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Моделировать"
        .tag = "FireModel"
        .TooltipText = "Моделировать развитие пожара тактической моделью"
        .FaceID = 896
        .BeginGroup = False
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
'---Процедура удаления кнопки "Площадь" из панели управления "Моделирование"--------------
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error Resume Next

    Set Bar = Application.CommandBars("Моделирование")
'---Удаление кнопки "Расчетная зона" из панели управления "Моделирование"------------------------
    Set Button = Bar.Controls("Расчетная зона")
    Button.Delete
'---Удаление кнопки "Моделировать" из панели управления "Моделирование"------------------------
    Set Button = Bar.Controls("Моделировать")
    Button.Delete

    
    
Set Button = Nothing
Set Bar = Nothing

End Sub
