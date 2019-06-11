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
'��������� ��� �������� ���������

'---������� ���������� ����������
Set vdAppEventsTech2 = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'��������� ��� �������� ���������

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---��������� ���������� ���������� ��� ����������� ������������ �� ��������� ����������� �����
    Set vdAppEventsTech2 = Visio.Application
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub vdAppEventsTech2_CellChanged(ByVal cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer

    On Error GoTo EX
'---��������� ��� ������
'MsgBox Cell.Name
    If cell.Name = "Prop.Set" Or cell.Name = "Prop.Model" Then
        '---��������� ��������� ��������� ������� �������
        ShpInd = cell.Shape.ID
        ModelsListImport (ShpInd)
        '---��������� ��������� ��� - �������
        GetTTH (ShpInd)


    End If

'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "vdAppEventsTech2_CellChanged"
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

End Sub
