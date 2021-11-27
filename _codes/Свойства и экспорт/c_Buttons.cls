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
Public WithEvents ComButReport As Office.CommandBarButton        'Донесение
Attribute ComButReport.VB_VarHelpID = -1
Public WithEvents ComButKBD As Office.CommandBarButton           'Карточка боевых действий
Attribute ComButKBD.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButReport = Application.CommandBars("Экспорт").Controls("Донесение")
    Set ComButKBD = Application.CommandBars("Экспорт").Controls("КБД")
End Sub

Private Sub Class_Terminate()
    Set ComButReport = Nothing
    Set ComButKBD = Nothing
End Sub


Private Sub ComButReport_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ExportToWord_Donesenie
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComButKBD_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ExportToWord_KBD
Tail:
    '---Выходим из процедуры обработки
End Sub


