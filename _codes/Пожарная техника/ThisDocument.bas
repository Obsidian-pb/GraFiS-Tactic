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
'��������� ��� �������� ���������
'---������� ���������� ����������
Set vdAppEvents = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'��������� ��� �������� ���������

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---��������� ���������� ���������� ��� ����������� ������������ �� ��������� ����������� �����
    Set vdAppEvents = Visio.Application

'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
'        MsgBox "�� �������� �����!"
    End If
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

End Sub

Public Sub DirectApp()
Set vdAppEvents = Visio.Application
End Sub

Private Sub vdAppEvents_CellChanged(ByVal cell As IVCell)
'��������� ���������� ������� � �������
Dim ShpInd As Integer

'---���������, ����� �� ������ ������������ ������ (��������, �������� ������)
    If cell.Shape Is Nothing Then Exit Sub

'---��������� ��� ������
'!!!�� ����� �����������!!!
'    If Not IsShapeLinkedToDataAndDropFirst(cell.Shape) Then
    If Not IsShapeLinked(cell.Shape) Then
'    If cell.Shape.CellExists("User.InPage", 0) = False Then
        If cell.Name = "Prop.Set" Then
'            Debug.Print cell.Name & " -> " & cell.Shape.Name
            '---��������� ��������� ��������� ������� �������
'            ShpInd = cell.Shape.ID
            ModelsListImport cell.Shape
        ElseIf cell.Name = "Prop.Model" Then
            '---��������� ��������� ���
'            ShpInd = cell.Shape.ID
            GetTTH cell.Shape
        End If
'    Else
''        cell.Shape.Cells("Prop.Set").FormulaU = "INDEX(0,Prop.Set.Format)"
''        ShapeLinkRefresh cell.Shape
    End If
    
'� ������, ���� ��������� ��������� �� ������ ������ ���������� �������
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
