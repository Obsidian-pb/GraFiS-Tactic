VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����� ��� ��������� ������� �� ������
Public WithEvents ComButReport As Office.CommandBarButton        '���������
Attribute ComButReport.VB_VarHelpID = -1
Public WithEvents ComButKBD As Office.CommandBarButton           '�������� ������ ��������
Attribute ComButKBD.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButReport = Application.CommandBars("�������").Controls("���������")
    Set ComButKBD = Application.CommandBars("�������").Controls("���")
End Sub

Private Sub Class_Terminate()
    Set ComButReport = Nothing
    Set ComButKBD = Nothing
End Sub


Private Sub ComButReport_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ExportToWord_Donesenie
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButKBD_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ExportToWord_KBD
Tail:
    '---������� �� ��������� ���������
End Sub


