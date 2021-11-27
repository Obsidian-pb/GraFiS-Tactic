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

Private WithEvents app  As Visio.Application
Attribute app.VB_VarHelpID = -1



Public Sub tempAct()
    Set app = Visio.Application
End Sub

Private Sub app_ConnectionsAdded(ByVal Connects As IVConnects)
    TurnIntoFormulaConnection Connects
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    '������������ ������ ������������
    DeleteButtons_f
    RemoveTB_f
    
    '������������ ������ ����������
    Set app = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    '���������, ����� �� ��������� "������"
    CheckReportsStencil
    
    '���������� ������ ������������
    AddTB_f
    AddButtons_f
    
    '���������� ������ ����������
    Set app = Visio.Application
End Sub



Public Sub CheckReportsStencil()
'��������� ��������� �� ��� �������� "������"
Const rep = "������.vss"
Dim stenc As Visio.Document
    
    On Error GoTo ex
    
    For Each stenc In Application.Documents
        If stenc.Name = rep Then
            'stenc.
            Exit Sub
        End If
    Next stenc
    
    Application.Documents.Open (ThisDocument.path & rep)
ex:
'������� ��� ��������
SaveLog Err, "CheckReportsStencil", "�������� ������� ������������� ��������� ������.vss"
End Sub


