VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Private WithEvents app As Visio.Application
Private sequencer As c_Sequencer
Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1




Private Sub app_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
    If KeyCode = 17 Then ctrlOn = True
End Sub
Private Sub app_KeyUp(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
    If KeyCode = 17 Then ctrlOn = False
End Sub




Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).visible = True
    
'---Обновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---Показываем панель управления РТП
    AddTB
    
'---Активируем автоматизатор временных свойств фигур ГраФиС
    ActivateSequencer
    
'---Получаем ссылку на приложение
    Set app = Visio.Application
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'---Деактивируем автоматизатор временных свойств фигур ГраФиС
    DeActivateSequencer
'---Скрываем панель управления РТП
    RemoveTB
'---Удаляем ссылку на приложение
    Set app = Nothing
End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If

End Sub








'Активация и деактивация контроллера временной последовательности фигур
Public Sub ActivateSequencer()
    Set sequencer = New c_Sequencer
End Sub
Public Sub DeActivateSequencer()
    Set sequencer = Nothing
End Sub


