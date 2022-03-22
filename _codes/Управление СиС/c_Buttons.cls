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

Public WithEvents ComButDescView As Office.CommandBarButton          'Показать описание БД
Attribute ComButDescView.VB_VarHelpID = -1
Public WithEvents ComButTechView As Office.CommandBarButton          'Показать список техники
Attribute ComButTechView.VB_VarHelpID = -1
Public WithEvents ComButPersonnel As Office.CommandBarButton         'Показать перечень личного состава и боевых позиций
Attribute ComButPersonnel.VB_VarHelpID = -1
Public WithEvents ComButNozzlesView As Office.CommandBarButton       'Показать список стволов
Attribute ComButNozzlesView.VB_VarHelpID = -1
Public WithEvents ComButGDZSView As Office.CommandBarButton          'Показать список ГДЗС
Attribute ComButGDZSView.VB_VarHelpID = -1
Public WithEvents ComButTimelineView As Office.CommandBarButton      'Показать таймлайн
Attribute ComButTimelineView.VB_VarHelpID = -1
Public WithEvents ComButStatistsView As Office.CommandBarButton      'Показать список статистов
Attribute ComButStatistsView.VB_VarHelpID = -1
Public WithEvents ComButExplicationView As Office.CommandBarButton   'Показать экспликацию
Attribute ComButExplicationView.VB_VarHelpID = -1

Public WithEvents ComButDescExport As Office.CommandBarButton        'Экспорт описания БД
Attribute ComButDescExport.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCommand = Application.CommandBars("РТП").Controls("Команда")
    Set ComButInfo = Application.CommandBars("РТП").Controls("Информация")
    Set ComButMark = Application.CommandBars("РТП").Controls("Оценка")
    
    Set ComButDescView = Application.CommandBars("РТП").Controls("Показать описание БД")
    Set ComButTechView = Application.CommandBars("РТП").Controls("Список техники")
    Set ComButPersonnel = Application.CommandBars("РТП").Controls("Список личного состава")
    Set ComButNozzlesView = Application.CommandBars("РТП").Controls("Список стволов")
    Set ComButGDZSView = Application.CommandBars("РТП").Controls("Список ГДЗС")
    Set ComButTimelineView = Application.CommandBars("РТП").Controls("Таймлайн")
'    Set ComButStatistsView = Application.CommandBars("РТП").Controls("Список статистов")
    Set ComButExplicationView = Application.CommandBars("РТП").Controls("Экспликация")
    
    Set ComButDescExport = Application.CommandBars("РТП").Controls("Экспорт описания БД")
End Sub

Private Sub Class_Terminate()
    Set ComButCommand = Nothing
    Set ComButInfo = Nothing
    Set ComButMark = Nothing
    
    Set ComButDescView = Nothing
    Set ComButTechView = Nothing
    Set ComButPersonnel = Nothing
    Set ComButNozzlesView = Nothing
    Set ComButGDZSView = Nothing
    Set ComButTimelineView = Nothing
    Set ComButStatistsView = Nothing
    Set ComButExplicationView = Nothing
    
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



'-------------------------Кнопки списков--------------------------------------------------
Private Sub ComButDescView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    DescriptionViewInList
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButTechView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowUnits
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButPersonnel_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowPersonnel
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButNozzlesView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowNozzles
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButGDZSView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowGDZS
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButTimelineView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowTimeLine
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButStatistsView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowStatists
Tail:
    '---Выходим из процедуры обработки
End Sub
Private Sub ComButExplicationView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки
    
'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Запускаем основное событие
    ShowExplication
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
