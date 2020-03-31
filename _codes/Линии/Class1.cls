VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents comBut As Office.CommandBarButton   'Рабочая линия
Attribute comBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  'Всасывающая линия
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  'Магистральные линии
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  'Проброс линии
Attribute ComBut4.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set comBut = Application.CommandBars("Превращения").Controls("Рукав")
    Set ComBut2 = Application.CommandBars("Превращения").Controls("Всасывающий рукав")
    Set ComBut3 = Application.CommandBars("Превращения").Controls("Магистральная линия")
    Set ComBut4 = Application.CommandBars("Превращения").Controls("Проброс")
End Sub

Private Sub Class_Terminate()
    Set comBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Обратить в рукавную линию"
    
    On Error GoTo EX

'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру линии
                MakeHoseLine Application.ActiveWindow.Selection(1), 51, 0
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
    
'            '---Обращаем фигуру в фигуру зону горения
'                MorphToLake

Exit Sub
EX:
    SaveLog Err, "ComBut_Click"
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие клика по кнопке "Всасывающая линия"

    On Error GoTo EX
    
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру всасывающей линии
                MakeVHoseLine
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
EX:
    SaveLog Err, "ComBut2_Click"
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие клика по кнопке "Магистральная линия"

    On Error GoTo EX
        
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру магистральной линии
                MakeHoseLine Application.ActiveWindow.Selection(1), 77, 1
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If

Exit Sub
EX:
    SaveLog Err, "ComBut3_Click"
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие клика по кнопке "Проброс"
    
    On Error GoTo EX
    
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---Обращаем фигуру в фигуру проброса
                MsgBox "Обращаем в проброс"
'                MakeHoseLine Application.ActiveWindow.Selection(1), 77, 1
                MakeHoseDrop Application.ActiveWindow.Selection(1), 77, 1
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
Exit Sub
EX:
    SaveLog Err, "ComBut4_Click"
End Sub

Public Sub PictureRefresh()
'Обновляем картинки на кнопках
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---Обновляем иконки на кнопках
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("Превращения").Controls("Рукав")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Hose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Hose2.bmp")
    End With
    
    Set cb = Application.CommandBars("Превращения").Controls("Всасывающий рукав")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\VHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\VHose2.bmp")
    End With
    
    Set cb = Application.CommandBars("Превращения").Controls("Магистральная линия")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\MHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MHose2.bmp")
    End With
    
    Set cb = Nothing
End Sub



'--------------------------------Внутренние функции класса---------------------
Private Function GetSelectedButtonCaption() As String
Dim v_Ctrl As CommandBarControl

    On Error GoTo EX
'---Проверяем какая кнопка нажата и в зависимости от этого выполняем действие
    For Each v_Ctrl In Application.CommandBars("Превращения").Controls
        If v_Ctrl.State = msoButtonDown Then
            GetSelectedButtonCaption = v_Ctrl.Caption
            Exit Function
        End If
    Next v_Ctrl
GetSelectedButtonCaption = ""

Exit Function
EX:
End Function



