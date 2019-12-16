VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ButEvent As c_Buttons

Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1



Private Sub app_KeyDown(ByVal KeyCode As Long, ByVal KeyButtonState As Long, CancelDefault As Boolean)
    MsgBox KeyCode
End Sub





Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'������������ �������� ���������
    On Error GoTo EX
    
'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
'---������� ������ ���������� "������" � ��������� �� ��� ������
'    AddTimer
    
'---������� ������ ���������� "�����������" � ��������� �� ��� ������
    AddTB_SpecFunc
    AddButtons

'---���������� ������ ������������ ������� ������
    Set ButEvent = New c_Buttons

'---��������� ������� ����������
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
EX:
   
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'������������ �������� ���������

'---������������ ������ ������������ ������� ������
    Set ButEvent = Nothing
    
'---������� ������ � ������ ���������� "�����������"
    DeleteButtons
    
'---������� ������ ���������� "������"
    DelTBTimer

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


'Public Sub SetApp()
'    Set app = Application
'End Sub
'
'
'Public Sub DelApp()
'    Set app = Nothing
'End Sub


