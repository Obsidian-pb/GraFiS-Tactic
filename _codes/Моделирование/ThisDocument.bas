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

Dim WithEvents SquareAppEvents As Visio.Application
Attribute SquareAppEvents.VB_VarHelpID = -1

Dim ButEventCalcArea As ClassCalcArea2
Dim ButEventModelling As ClassMod2

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo EX
    StartAction

Exit Sub
EX:
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    EndAction
End Sub


Public Sub StartAction()
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    

'---Добавляем ячейку Аспект (если еще не была добавлена)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---Инициируем объект SquareAppEvents для реагирования на действия пользователя
    Set SquareAppEvents = Visio.Application
    
'---Создаем панель управления "Превращения" и добавляем на нее кнопку "Обратить в зону горения"
    AddTBImagination
    AddButtons
    
'---Активируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Включаем показ окон
    VfB_NotShowPropertiesWindow = False


    Set ButEventCalcArea = New ClassCalcArea2
    Set ButEventModelling = New ClassMod2
End Sub

Public Sub EndAction()
'Процедура закрытия документа и удаления его рабочих элементов

'---Очищаем объект ButEvent и удаляем кнопку "Площадь" с панели управления "Моделирование"
    Set ButEventCalcArea = Nothing
    Set ButEventModelling = Nothing
    DeleteButtons
    
'---Деактивируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then Set app = Nothing
    
'---В случае, если на панели "Моделирование нет ни одной кнопки, удаляем её
    If Application.CommandBars("Моделирование").Controls.Count = 0 Then RemoveTBImagination
'---Очищаем переменную приложения
    Set SquareAppEvents = Nothing
End Sub

