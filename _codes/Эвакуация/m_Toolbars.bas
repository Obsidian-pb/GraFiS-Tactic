Attribute VB_Name = "m_Toolbars"
Option Explicit



Public Sub AddTB_Evacuation()
'Процедура добавления панели управления "Эвакуация"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim i As Integer
    
'---Проверяем есть ли уже панель управления "Эвакуация"------------------------------
    For i = 1 To Application.CommandBars.count
        If Application.CommandBars(i).Name = "Эвакуация" Then Exit Sub
    Next i

'---Создаем панель управления "Эвакуация"--------------------------------------------
    With Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
        .Name = "Эвакуация"
        .Visible = True
    End With
    
    AddButtons
End Sub

Public Sub RemoveTB_Evacuation()
'Процедура добавления панели управления "Эвакуация"-------------------------------
    On Error GoTo EX
    Application.CommandBars("Эвакуация").Delete

EX:
End Sub

Public Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "Эвакуация"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Эвакуация")
    DocPath = ThisDocument.path
    
'---Добавляем кнопки на панель управления "Эвакуация"--------------------------------
'---Кнопка "Выбрать все фигуры графа"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Выбрать все"
        .Tag = "SelectAllGraphShapes"
        .TooltipText = "Выбрать все фигуры графа"
        .FaceID = 1446
    End With
'---Кнопка "Перенумеровать все фигуры графа"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Перенумеровать"
        .Tag = "Renum"
        .TooltipText = "Перенумеровать все фигуры графа"
        .FaceID = 1116
    End With

'---Кнопка "Рассчитать время эвакуации"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Рассчитать"
        .Tag = "Calculate"
        .TooltipText = "Рассчитать время эвакуации"
        .FaceID = 283
        .BeginGroup = True
    End With
    

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub

