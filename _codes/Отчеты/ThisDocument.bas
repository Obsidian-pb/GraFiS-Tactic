VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vmO_App As Visio.Application
Attribute vmO_App.VB_VarHelpID = -1





Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    
    On Error Resume Next
    
'---Деактивируем объект отслеживания нажатия кнопок
    Set ButEvent = Nothing
    
'---Закрываем окно "Мастер проверок"
    TacticDataForm.CloseThis
    WarningsForm.CloseThis
    
'---Удаляем кнопки с панели управления "СпецФункции"
    DeleteButtons
    
'---Уничтожаем InfoCollector
    KillA
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
Set vmO_App = Visio.Application
    
'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
   
'---Добавляем панель инструментов "Спецфункции"
    AddTB_SpecFunc
    AddButtons
    
'---Активируем и показываем окно "Мастер проверок" и запускаем предварительный анализ (Включить для показа формы по-умолчанию)
'    MCheckForm.Show
'    MasterCheckRefresh
       
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates
    
End Sub



Private Sub vmO_App_CellChanged(ByVal Cell As IVCell)
'Процедура проверят была ли изменена ячейка Prop.FireMax или Prop.TimeMax и запускает процедуру обновления графика
Dim vsS_CellName As String

    vsS_CellName = Cell.Name
'    If vsS_CellName = "Prop.FireMax" Or vsS_CellName = "Prop.TimeMax" Then
'        sP_ChangeGraphDirect (Cell.Shape.ID)
'    End If
'    If vsS_CellName = "Prop.FireMax" Or vsS_CellName = "Prop.TimeMax" Or vsS_CellName = "Prop.WaterIntenseH" _
'                                                                Or vsS_CellName = "Prop.WaterIntenseType" Then
'        sP_ChangeGraphDirect (Cell.Shape.ID)
'    End If   ' С ИНТЕНСИВНОСТЬЮ !!!!!!!

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

