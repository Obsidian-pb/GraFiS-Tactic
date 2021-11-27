VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vdAppEvents As Visio.Application
Attribute vdAppEvents.VB_VarHelpID = -1

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура при закрытии документа
'---Очищаем переменную приложения
Set vdAppEvents = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'Процедура при открытии документа

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set vdAppEvents = Visio.Application

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
'        MsgBox "Не Цветовая схема!"
    End If
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

End Sub

Public Sub DirectApp()
Set vdAppEvents = Visio.Application
End Sub

Private Sub vdAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer

'---Проверяем, имеет ли ячейка родительскую фигуру (например, является стилем)
    If cell.Shape Is Nothing Then Exit Sub

'---Проверяем имя ячейки
'!!!До конца разобраться!!!
'    If Not IsShapeLinkedToDataAndDropFirst(cell.Shape) Then
    If Not IsShapeLinked(cell.Shape) Then
'    If cell.Shape.CellExists("User.InPage", 0) = False Then
        If cell.Name = "Prop.Set" Then
'            Debug.Print cell.Name & " -> " & cell.Shape.Name
            '---Запускаем процедуру получения списков моделей
'            ShpInd = cell.Shape.ID
            ModelsListImport cell.Shape
        ElseIf cell.Name = "Prop.Model" Then
            '---Процедура получения ТТХ
'            ShpInd = cell.Shape.ID
            GetTTH cell.Shape
        End If
'    Else
''        cell.Shape.Cells("Prop.Set").FormulaU = "INDEX(0,Prop.Set.Format)"
''        ShapeLinkRefresh cell.Shape
    End If
    
'В случае, если произошло изменение не нужной ячейки прекращаем событие
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
