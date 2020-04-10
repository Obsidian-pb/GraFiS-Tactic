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

Dim WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1



Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    '������� ������ �� ����������
'    Set visApp = Visio.Application
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo Tail
    
'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---����������� ������ visApp � ������ �� ���������� Visio
'    Set visApp = Visio.Application

'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "��������� ������� ������! ���� ��� ����� �����������, ��������� � �������������."
End Sub






