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

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 1000

Dim ButEvent As c_Buttons









Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'Обрабатываем открытие документа
    On Error GoTo ex
    
'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
'---Создаем панель управления "Спецфункции" и добавляем на нее кнопки
    AddTB_SpecFunc
    AddButtons

'---Активируем объект отслеживания нажатия кнопок
    Set ButEvent = New c_Buttons

'---Активируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

''---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates

Exit Sub
ex:
   
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Обрабатываем закрытие документа

'---Деактивируем объект отслеживания нажатия кнопок
    Set ButEvent = Nothing
    
'---Удаляем кнопки с панели управления "СпецФункции"
    DeleteButtons
    
'---Деактивируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then Set app = Nothing
    
'---Удаляем панель управления "Таймер"
    DelTBTimer

End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim Cell As Visio.Cell

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


Private Sub app_CellChanged(ByVal Cell As IVCell)
'---Один раз в выполняем обновление иконок на кнопках
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub


