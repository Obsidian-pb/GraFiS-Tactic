VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut4 As Office.CommandBarButton
Attribute ComBut4.VB_VarHelpID = -1


Private Sub Class_Initialize()
'    MsgBox "123"
    Set ComBut4 = Application.CommandBars("Превращения").Controls("Обрушение")
End Sub

Private Sub Class_Terminate()
    Set ComBut4 = Nothing
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

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
            '---Обращаем в зону обрушения
                MorphToRush
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

Public Sub MorphToRush()

    On Error GoTo EX
'---Проверяем имеется ли возможность обращения фигуры
'    If Not CheckRushShape Then
'    '---Отмечаем кнопку
'        PS_CheckButtons ComBut4
'        Exit Sub
'    End If

'---Запускаем процедуру обращения выбранной фигуры в зону обрушения
    ImportRushInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToRush"
End Sub


