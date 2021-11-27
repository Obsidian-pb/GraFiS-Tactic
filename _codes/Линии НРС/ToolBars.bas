Attribute VB_Name = "ToolBars"

Sub AddTBImagination()
'Процедура добавления панели управления "Превращения"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    'Const DocPath = ThisDocument.Path
    Dim DocPath As String
    
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
    On Error Resume Next
    
    Application.CommandBars("Превращения").Delete
End Sub



Sub AddTBNRSCalc()
'Процедура добавления панели управления "Насосно-рукавные системы"-------------------------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
'---Проверяем есть ли уже панель управления "Насосно-рукавные системы"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "Насосно-рукавные системы" Then Exit Sub
    Next i

'---Создаем панель управления "Превращения"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "Насосно-рукавные системы"
        .Visible = True
    End With

End Sub

Sub RemoveTBNRSCalc()
'Процедура добавления панели управления "Насосно-рукавные системы"-------------------------------
    On Error Resume Next
    
    Application.CommandBars("Насосно-рукавные системы").Delete
End Sub

'--------------------------------------Кнопка рабочая линия-------------------------
Sub AddButtonLine()
'Процедура добавление новой кнопки на панель управления "Превращения"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
    DocPath = ThisDocument.path

'---Проверяем есть ли уже на панели управления "Превращения" кнопка "Рукав"------------------------------
'    For i = 1 To Application.CommandBars("Превращения").Controls.Count
''    Application.CommandBars("Превращения").Controls
'        If Application         'CommandBars("Превращения").Controls(i).Name = "Рукав" Then

'    Next i
    
'    If Not Application.CommandBars("Превращения").Controls("Рукав") = Nothing Then
'        Set Bar = Nothing
'        Exit Sub
'    End If
    
'---Добавляем кнопки на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить в рукавную линию"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Рукав"
        .Tag = "Hose"
        '.OnAction = "Application.Documents('Линии НРС.vss').ExecuteLine ('ProvExchange')"
        .TooltipText = "Обратить в рабочую рукавную линию"
        .Picture = LoadPicture(DocPath & "Bitmaps\Hose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Hose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub


Sub DeleteButtonLine()
'---Процедура удаления кнопки "Рукав" с панели управления "Превращения"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
'---Удаление кнопки "Рукав" на панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Рукав")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------Кнопка магистральная линия-------------------------
Sub AddButtonMLine()
'Процедура добавление новой кнопки "Магистральная линия" на панель управления "Превращения"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
    DocPath = ThisDocument.path
    
'---Добавляем кнопку на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить в магистральную рукавную линию"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Магистральная линия"
        .Tag = "MHose"
        .TooltipText = "Обратить в магистральную рукавную линию"
        .Picture = LoadPicture(DocPath & "Bitmaps\MHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MHose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonMLine()
'---Процедура удаления кнопки "Магистральная линия" с панели управления "Превращения"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
'---Удаление кнопки "Рукав" на панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Магистральная линия")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub


'--------------------------------------Кнопка всасывающая линия-------------------------
Sub AddButtonVHose()
'Процедура добавление новой кнопки на панель управления "Превращения"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
    DocPath = ThisDocument.path
'---Добавляем кнопки на панель управления "Превращения"--------------------------------
'---Кнопка "Обратить во всасывающую линию"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Всасывающий рукав"
        .Tag = "VHose"
        .TooltipText = "Обратить во всасывающую рукавную линию"
        .Picture = LoadPicture(DocPath & "Bitmaps\VHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\VHose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonVHose()
'Процедура удаления кнопки "всасывающий рукав" с панели управления "Превращения"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Превращения")
'---Удаление кнопки "Рукав" на панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Всасывающий рукав")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------Кнопка Расчет НРС-------------------------
Sub AddButtonNormalize()
'Процедура добавление новой кнопки на панель управления "Насосно-рукавные системы"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
    DocPath = ThisDocument.path
'---Добавляем кнопки на панель управления "Насосно-рукавные системы"--------------------------------
'---Кнопка "Нормализация"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Расчет НРС"
        .Tag = "Calculate NRS"
        .TooltipText = "Рассчитать НРС"
        .FaceID = 807
        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNormalize()
'Процедура удаления кнопки "Нормализация" с панели управления "Насосно-рукавные системы"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
'---Удаление кнопки "Нормализация" на панели управления "Насосно-рукавные системы"------------------------
    Set Button = Bar.Controls("Расчет НРС")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------Кнопка Настройка расчета НРС-------------------------
Sub AddButtonNRSSettings()
'Процедура добавление новой кнопки на панель управления "Насосно-рукавные системы"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
    DocPath = ThisDocument.path
'---Добавляем кнопки на панель управления "Насосно-рукавные системы"--------------------------------
'---Кнопка "Настройка расчета НРС"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Настройка расчета НРС"
        .Tag = "NRS Settings"
        .TooltipText = "Настройки расчета НРС"
        .FaceID = 642
'        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNRSSettings()
'Процедура удаления кнопки "Настройка расчета НРС" с панели управления "Насосно-рукавные системы"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
'---Удаление кнопки "Нормализация" на панели управления "Превращения"------------------------
    Set Button = Bar.Controls("Настройка расчета НРС")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------Кнопка Отчет расчета НРС-------------------------
Sub AddButtonNRSReport()
'Процедура добавление новой кнопки на панель управления "Насосно-рукавные системы"--------------

'---Объявляем переменные и постоянные--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
    DocPath = ThisDocument.path
'---Добавляем кнопки на панель управления "Насосно-рукавные системы"--------------------------------
'---Кнопка "Настройка расчета НРС"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Отчет расчета НРС"
        .Tag = "NRS Report"
        .TooltipText = "Показать отчет расчета НРС"
        .FaceID = 230
'        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNRSReport()
'Процедура удаления кнопки "Отчет расчета НРС" с панели управления "Насосно-рукавные системы"--------------
    
    On Error Resume Next
    
'---Объявляем переменные и постоянные-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("Насосно-рукавные системы")
'---Удаление кнопки "Отчет расчета НРС" на панели управления "Насосно-рукавные системы"------------------------
    Set Button = Bar.Controls("Отчет расчета НРС")
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
