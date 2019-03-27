VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Класс для обработки нажатий на кнопки
Public WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1

Public WithEvents ConstructionComBut As Office.CommandBarButton   'Маска
Attribute ConstructionComBut.VB_VarHelpID = -1
Public WithEvents ConstructionComBut2 As Office.CommandBarButton  'Рисование стен
Attribute ConstructionComBut2.VB_VarHelpID = -1

Private wallDrawer As c_WallDrawer

Private maskClass As c_Mask


Private Sub Class_Initialize()
    Set ConstructionComBut = Application.CommandBars("Конструкции").Controls("Маска")
    Set ConstructionComBut2 = Application.CommandBars("Конструкции").Controls("Рисование стен")
    
    Set maskClass = New c_Mask
End Sub

Private Sub Class_Terminate()
    Set ConstructionComBut = Nothing
    Set ConstructionComBut2 = Nothing
    
    Set maskClass = Nothing
End Sub

Private Sub ConstructionComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Наложение маски"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    maskClass.s_makeMask
    
Exit Sub
Tail:
    '---Выходим из процедуры обработки
    MsgBox Err.description
End Sub

Private Sub ConstructionComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Рисовальщик стен"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Проверяем состояние кнопки
    If Ctrl.State = msoButtonDown Then
        wallDrawer.Term
        Set wallDrawer = Nothing
        Set visApp = Nothing
    Else
    '---Активируем рисовальщик
        Set wallDrawer = New c_WallDrawer
    '---Включаем кнопку, если рисовальщик активировался
        If wallDrawer.Init Then Ctrl.State = msoButtonDown
    '---Активируем класс слежения за приложением
        Set visApp = Application
    End If

Exit Sub
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub visApp_BeforeShapeDelete(ByVal Shape As IVShape)
'---Проверяем не была ли удалена фигура стены выбранная в качестве мастера для рисовальщика
    If Shape = wallDrawer.wallShape Then
        wallDrawer.Term
        Set wallDrawer = Nothing
        Set visApp = Nothing
    End If
End Sub
