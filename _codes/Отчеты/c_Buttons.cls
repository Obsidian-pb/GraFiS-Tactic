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

Public WithEvents ComButCheck As Office.CommandBarButton        'Мастер проверок
Attribute ComButCheck.VB_VarHelpID = -1
Public WithEvents ComButTactic As Office.CommandBarButton       'Тактические данные
Attribute ComButTactic.VB_VarHelpID = -1
Public WithEvents ComButExportToWord As Office.CommandBarButton 'Экспорт в Word
Attribute ComButExportToWord.VB_VarHelpID = -1
Public WithEvents ComButSettings As Office.CommandBarButton     'Настройки
Attribute ComButSettings.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCheck = Application.CommandBars("Спецфункции").Controls("Мастер проверок")
    Set ComButTactic = Application.CommandBars("Спецфункции").Controls("Тактические данные")
    Set ComButExportToWord = Application.CommandBars("Спецфункции").Controls("Экспорт в Word")
    Set ComButSettings = Application.CommandBars("Спецфункции").Controls("Настройки")
End Sub

Private Sub Class_Terminate()
    Set ComButCheck = Nothing
    Set ComButTactic = Nothing
    Set ComButExportToWord = Nothing
    Set ComButSettings = Nothing
End Sub


Private Sub ComButCheck_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Проверяем показана ли уже форма
    If WarningsForm.Visible = True Then Exit Sub
    
'---Запускаем основное событие
    WarningsForm.Activate.Refresh
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComButTactic_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
       
'---Проверяем показана ли уже форма
    If TacticDataForm.Visible = True Then Exit Sub
    
'---Запускаем основное событие
    TacticDataForm.Activate.Refresh
    
''---Показываем кнопку "Экспорт в Word"
'    Application.CommandBars("Спецфункции").Controls("Экспорт в Word").Visible = True
    
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComButExportToWord_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки "Экспорт в Word" (Пока только тактические данные)

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Проверяем показана ли уже форма Тактических данных
    If TacticDataForm.Visible = False Then Exit Sub
    
'---Запускаем основное событие
    TacticDataForm.ExportToWord
Tail:
    '---Выходим из процедуры обработки
End Sub

Private Sub ComButSettings_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Отслеживаем событие нажатия кнопки

'---Включаем обработку ошибок - для предотвращения выброса класса при попытке обращения ничего
    On Error GoTo Tail
    
'---Проверяем показана ли уже форма
    If CommOptForm.Visible = True Then Exit Sub
    
'---Запускаем основное событие
    CommOptForm.ShowForm
Tail:
    '---Выходим из процедуры обработки
End Sub


