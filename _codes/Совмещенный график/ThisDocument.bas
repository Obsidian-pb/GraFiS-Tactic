VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
''---��������� ������� ����������
'    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub AddTimeUserCells()
'����� ��������� ������ "User.FireTime", "User.CurrentTime"
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
    If Not docSheet.CellExists("User.FireEndTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireEndTime", visTagDefault
        docSheet.Cells("User.FireEndTime").FormulaU = "User.FireTime+0.05"
    End If

End Sub

