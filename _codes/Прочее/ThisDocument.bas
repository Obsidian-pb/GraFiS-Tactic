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

    On Error GoTo Tail
    
'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������."
End Sub


