VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassStorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut2 As Office.CommandBarButton
Attribute ComBut2.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut2 = Application.CommandBars("Превращения").Controls("Шторм")
End Sub

Private Sub Class_Terminate()
    Set ComBut2 = Nothing
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    On Error GoTo EX
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.state = msoButtonDown Then
        Ctrl.state = msoButtonUp
        Exit Sub
    End If

    If Ctrl.state = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If IsHavingUserSection(True) And IsSquare(True) Then
            '---Обращаем в огненный шторм
                MorphToStorm
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

Public Sub MorphToStorm()

    On Error GoTo EX
    
'---Запускаем процедуру обращения выбранной фигуры в огненный шторм
    ImportStormInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToStorm"
End Sub

Public Sub PictureRefresh()
'Обновляем картинки на кнопках
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---Обновляем иконки на кнопках
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("Превращения").Controls("Шторм")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Storm1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Storm2.bmp")
    End With
    
Set cb = Nothing
End Sub
