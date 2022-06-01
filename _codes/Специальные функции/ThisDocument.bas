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

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 1000

Dim ButEvent As c_Buttons









Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'������������ �������� ���������
    On Error GoTo ex
    
'---��������� ������ "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
'---������� ������ ���������� "�����������" � ��������� �� ��� ������
    AddTB_SpecFunc
    AddButtons

'---���������� ������ ������������ ������� ������
    Set ButEvent = New c_Buttons

'---���������� ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

''---��������� ������� ����������
'    fmsgCheckNewVersion.CheckUpdates

Exit Sub
ex:
   
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'������������ �������� ���������

'---������������ ������ ������������ ������� ������
    Set ButEvent = Nothing
    
'---������� ������ � ������ ���������� "�����������"
    DeleteButtons
    
'---������������ ������ ������������ ��������� � ���������� ��� 201� ������
    If Application.version > 12 Then Set app = Nothing
    
'---������� ������ ���������� "������"
    DelTBTimer

End Sub

Private Sub AddTimeUserCells()
'����� ��������� ������ "User.FireTime", "User.CurrentTime"
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


Private Sub app_CellChanged(ByVal Cell As IVCell)
'---���� ��� � ��������� ���������� ������ �� �������
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub


