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

'Private WithEvents app As Visio.Application
Private sequencer As c_Sequencer
Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1




Private Sub app_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
    If KeyCode = 17 Then ctrlOn = True
End Sub
Private Sub app_KeyUp(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
    If KeyCode = 17 Then ctrlOn = False
End Sub




Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---���������� ���� �������
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).visible = True
    
'---���������/������������ � �������� �������� ����� ���������
    '---��������� �� �������� �� �������� �������� ���������� �������� �����
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---���������� ������ ���������� ���
    AddTB
    
'---���������� ������������� ��������� ������� ����� ������
    ActivateSequencer
    
'---�������� ������ �� ����������
    Set app = Visio.Application
    
'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'---������������ ������������� ��������� ������� ����� ������
    DeActivateSequencer
'---�������� ������ ���������� ���
    RemoveTB
'---������� ������ �� ����������
    Set app = Nothing
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








'��������� � ����������� ����������� ��������� ������������������ �����
Public Sub ActivateSequencer()
    Set sequencer = New c_Sequencer
End Sub
Public Sub DeActivateSequencer()
    Set sequencer = Nothing
End Sub


