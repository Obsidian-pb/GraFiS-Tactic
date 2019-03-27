VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFireArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton
Attribute ComBut.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("Превращения").Controls("Площадь")
'MsgBox ComBut.Caption
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
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
            '---Обращаем фигуру в фигуру зоны горения
                MorphToFireArea
            End If
        Else
'            MsgBox "Не выбрано ни одной фигуры! Будет включен режим автоматического обращения!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
EX:
    SaveLog Err, "ComBut_Click"
End Sub

Public Sub MorphToFireArea()

    On Error GoTo EX

'---Проверяем имеется ли возможность обращения фигуры
'    If Not CheckSquareShape Then
'    '---Отмечаем кнопку
'        PS_CheckButtons ComBut 'Ctrl
'        Exit Sub
'    End If

'---Запускаем процедуру обращения выбранной фигуры в площадь пожара
    ImportAreaInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToFireArea"
End Sub
