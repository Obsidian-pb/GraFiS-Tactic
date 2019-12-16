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
'Dim vOC_InfoAnalizer As InfoCollector




Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Set vOC_InfoAnalizer = Nothing

sP_InfoCollectorDeActivate 'Деактивируем класс InfoCollector через обращение к модулю m_Analize
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
Set vmO_App = Visio.Application

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Активируем класс InfoCollector через обращение к модулю m_Analize
    sP_InfoCollectorActivate
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates
    
End Sub



Private Sub vmO_App_CellChanged(ByVal cell As IVCell)
'Процедура проверят была ли изменена ячейка Prop.FireMax или Prop.TimeMax и запускает процедуру обновления графика
Dim vsS_CellName As String

    vsS_CellName = cell.Name
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
Sub show_m_chek_form()
MCheckForm.Show
End Sub
