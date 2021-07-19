Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTB()
'Процедура добавления панели управления "РТП"-------------------------------
Dim i As Integer
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---Проверяем есть ли уже панель управления "РТП"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "РТП" Then Exit Sub
    Next i

'---Создаем панель управления "РТП"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "РТП"
        .visible = True
    End With

'---Добавляем кнопки на панель управления
    AddButtons
End Sub

Sub RemoveTB()
'Процедура добавления панели управления "РТП"-------------------------------
    On Error Resume Next
    Application.CommandBars("РТП").Delete
End Sub

Sub AddButtons()
'Процедура добавление новой кнопки на панель управления "РТП"--------------
    
'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("РТП")
    
'---Добавляем кнопки на панель управления "РТП"--------------------------------
'---Кнопка "Команда"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Команда"
        .Tag = "Command"
        .TooltipText = "Команда тактической единице"
        .FaceID = 346
    End With
'---Кнопка "Информация"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Информация"
        .Tag = "Info"
        .TooltipText = "Информация для фигуры"
        .FaceID = 487 ' 162
        .beginGroup = True
    End With
'---Кнопка "Оценка"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Оценка"
        .Tag = "Mark"
        .TooltipText = "Оценка участника боевых действий или личного состава"
        .FaceID = 215 ' 162
    End With
    
'---Кнопки списков-------------------------------------------------
    '---"Описание"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Показать описание БД"
        .Tag = "DescriptionView"
        .TooltipText = "Показать описание БД"
        .beginGroup = True
        .FaceID = 5
'        .Caption = "Описание"
    End With
    '---"Техника"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Список техники"
        .Tag = "TechView"
        .TooltipText = "Показать список техники"
'        .BeginGroup = True
        .FaceID = 1277
    End With
    '---"Стволы"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Список стволов"
        .Tag = "NozzlesView"
        .TooltipText = "Показать список стволов"
'        .BeginGroup = True
        .FaceID = 2644
    End With
    '---"ГДЗС"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Список ГДЗС"
        .Tag = "GDZSView"
        .TooltipText = "Показать звенья и посты ГДЗС"
'        .BeginGroup = True
        .FaceID = 1253
    End With
    '---"Таймлайн"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Таймлайн"
        .Tag = "TimelineView"
        .TooltipText = "Показать таймлайн модели"
'        .BeginGroup = True
        .FaceID = 11
    End With
    '---"Статисты"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Список статистов"
        .Tag = "StatistsView"
        .TooltipText = "Показать сведения о статистах"
'        .BeginGroup = True
        .FaceID = 2141
    End With
    
    
    
'---Кнопка "Экспорт описания БД"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "Экспорт описания БД"
        .Tag = "DescriptionExport"
        .TooltipText = "Экспорт описания БД в Word"
        .beginGroup = True
        .FaceID = 582
    End With
    
    Set btns = New c_Buttons
    
    Set Bar = Nothing

Exit Sub
EX:
    Set Bar = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddButtons", "Команды"
End Sub


Sub DeleteButtons()
'---Процедура удаления кнопки "Команда" из панели управления "РТП"--------------
    On Error GoTo EX
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("РТП")
'---Удаление кнопки "Команда" из панели управления "РТП"------------------------
    Set Button = Bar.Controls("Команда")
    Button.Delete
'---Удаление кнопки "Информация" из панели управления "РТП"------------------------
    Set Button = Bar.Controls("Информация")
    Button.Delete
'---Удаление кнопки "Оценка" из панели управления "РТП"------------------------
    Set Button = Bar.Controls("Оценка")
    Button.Delete
'---Удаление кнопки "Экспорт описания БД" из панели управления "РТП"------------------------
    Set Button = Bar.Controls("Экспорт описания БД")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
EX:
'Выходим из процедуры
End Sub
