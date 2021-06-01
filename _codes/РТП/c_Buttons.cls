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
Public WithEvents ComButCommand As Office.CommandBarButton           'Команда
Attribute ComButCommand.VB_VarHelpID = -1
Public WithEvents ComButInfo As Office.CommandBarButton              'Информация
Attribute ComButInfo.VB_VarHelpID = -1
Public WithEvents ComButMark As Office.CommandBarButton              'Оценка
Attribute ComButMark.VB_VarHelpID = -1

Public WithEvents ComButDescExport As Office.CommandBarButton        'Экспорт описания БД
Attribute ComButDescExport.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCommand = Application.CommandBars("РТП").Controls("Команда")
    Set ComButInfo = Application.CommandBars("РТП").Controls("Информация")
    Set ComButMark = Application.CommandBars("РТП").Controls("Оценка")
    
    Set ComButDescExport = Application.CommandBars("РТП").Controls("Экспорт описания БД")
End Sub

Private Sub Class_Terminate()
    Set ComButCommand = Nothing
    Set ComButInfo = Nothing
    Set ComButMark = Nothing
    
    Set ComButDescExport = Nothing
End Sub


Private Sub ComButCommand_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    frm_Command.NewCommand
Tail:
    '---Выходим из процедуры обработки
End Sub


Private Sub ComButDescExport_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    DescriptionExportToWord
Tail:
    '---Выходим из процедуры обработки
End Sub

'-------------------------Кнопки информации--------------------------------------------------
Private Sub ComButInfo_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    frm_InfoSimple.NewInfo
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComButMark_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    frm_InfoSimple.NewInfo 1
Tail:
    '---Выходим из процедуры обработки
End Sub
