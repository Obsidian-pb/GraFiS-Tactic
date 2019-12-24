VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents comBut As Office.CommandBarButton
Attribute comBut.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set comBut = Application.CommandBars("Превращения").Controls("Естественный водоисточник")
'MsgBox ComBut.Caption
End Sub

Private Sub Class_Terminate()
    Set comBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    On Error GoTo EX
'---Проверяем включена ли кнопка - если включена - отключаем
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---Если выбрана хоть одна фигура - пытаемся ее обратить
            If IsHavingUserSection(True) And IsSquare(True) Then
            '---Обращаем фигуру в фигуру зону горения
                MorphToLake
            End If
        Else
            PS_CheckButtons Ctrl
        End If
    End If
    
    
Exit Sub
EX:
End Sub

Public Sub MorphToLake()

    On Error GoTo EX

'---Запускаем процедуру обращения выбранной фигуры в естественный водоисточник
    ImportOpenWaterInformation
    
Exit Sub
EX:
End Sub

Public Sub PictureRefresh()
'Обновляем картинки на кнопках
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---Обновляем иконки на кнопках
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("Превращения").Controls("Естественный водоисточник")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Lake1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Lake2.bmp")
    End With
    
Set cb = Nothing
End Sub

