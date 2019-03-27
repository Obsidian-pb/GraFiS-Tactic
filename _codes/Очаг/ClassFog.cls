VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut3 As Office.CommandBarButton
Attribute ComBut3.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut3 = Application.CommandBars("Превращения").Controls("Задымление")
End Sub

Private Sub Class_Terminate()
    Set ComBut3 = Nothing
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

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
            '---Обращаем фигуру в зону задымления
                MorphToFog
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

Public Sub MorphToFog()
'Обращаем фигуру в зону задымления

    On Error GoTo EX
'---Проверяем имеется ли возможность обращения фигуры
'    If Not CheckSquareShape Then
'    '---Отмечаем кнопку
'        PS_CheckButtons ComBut3
'        Exit Sub
'    End If
    
'---Запускаем процедуру обращения выбранной фигуры в зону задымления
    ImportFogInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToFog"
End Sub

