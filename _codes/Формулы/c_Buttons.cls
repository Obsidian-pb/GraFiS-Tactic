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
Public WithEvents ComButRefresh As Office.CommandBarButton        'Обновить
Attribute ComButRefresh.VB_VarHelpID = -1
Public WithEvents ComButShowAll As Office.CommandBarButton        'Показать все вычисления
Attribute ComButShowAll.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButRefresh = Application.CommandBars("Формулы").Controls("Обновить")
    Set ComButShowAll = Application.CommandBars("Формулы").Controls("Показать все")
End Sub

Private Sub Class_Terminate()
    Set ComButRefresh = Nothing
End Sub


Private Sub ComButRefresh_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса
    On Error GoTo Tail
    
'---Запускаем основное событие
    RefreshAllFormulas
Tail:
    '---Выходим из процедуры обработки
End Sub


Private Sub ComButShowAll_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowAllFormulas
Tail:
    '---Выходим из процедуры обработки
End Sub
