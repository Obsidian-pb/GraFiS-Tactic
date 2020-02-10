VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vdAppEventsTech2 As Visio.Application
Attribute vdAppEventsTech2.VB_VarHelpID = -1

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура при закрытии документа

'---Очищаем переменную приложения
Set vdAppEventsTech2 = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'Процедура при открытии документа

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---Обновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set vdAppEventsTech2 = Visio.Application
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub vdAppEventsTech2_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer

    On Error GoTo EX
'---Проверяем имя ячейки

    If cell.Name = "Prop.Set" Then
        '---Запускаем процедуру получения списков моделей
'        ShpInd = cell.Shape.ID
        ModelsListImport cell.Shape
    ElseIf cell.Name = "Prop.Model" Then
        '---Процедура получения ТТХ - СДЕЛАТЬ
        If Not IsShapeHaveCalloutAndDropFirst(cell.Shape) Then
'            ShpInd = cell.Shape.ID
            GetTTH cell.Shape
        End If
    End If

'В случае, если произошло изменение не нужной ячейки прекращаем событие
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "vdAppEventsTech2_CellChanged"
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
