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

Public WithEvents ComButCheck As Office.CommandBarButton        '������ ��������
Attribute ComButCheck.VB_VarHelpID = -1
Public WithEvents ComButTactic As Office.CommandBarButton       '����������� ������
Attribute ComButTactic.VB_VarHelpID = -1
'Public WithEvents ComButExportToWord As Office.CommandBarButton '������� � Word
Public WithEvents ComButSettings As Office.CommandBarButton     '���������
Attribute ComButSettings.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCheck = Application.CommandBars("�����������").Controls("������ ��������")
    Set ComButTactic = Application.CommandBars("�����������").Controls("����������� ������")
'    Set ComButExportToWord = Application.CommandBars("�����������").Controls("������� � Word")
    Set ComButSettings = Application.CommandBars("�����������").Controls("���������")
End Sub

Private Sub Class_Terminate()
    Set ComButCheck = Nothing
    Set ComButTactic = Nothing
'    Set ComButExportToWord = Nothing
    Set ComButSettings = Nothing
End Sub


Private Sub ComButCheck_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �� ��� �����
    If WarningsForm.Visible = True Then Exit Sub
    
'---��������� �������� �������
    WarningsForm.Activate.Refresh
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButTactic_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
       
'---��������� �������� �� ��� �����
    If TacticDataForm.Visible = True Then Exit Sub
    
'---��������� �������� �������
    TacticDataForm.Activate.Refresh
    
''---���������� ������ "������� � Word"
'    Application.CommandBars("�����������").Controls("������� � Word").Visible = True
    
Tail:
    '---������� �� ��������� ���������
End Sub

'Private Sub ComButExportToWord_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
''����������� ������� ������� ������ "������� � Word" (���� ������ ����������� ������)
'
''---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
'    On Error GoTo Tail
'
''---��������� �������� �� ��� ����� ����������� ������
'    If TacticDataForm.Visible = False Then Exit Sub
'
''---��������� �������� �������
'    TacticDataForm.ExportToWord
'Tail:
'    '---������� �� ��������� ���������
'End Sub

Private Sub ComButSettings_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �� ��� �����
    If CommOptForm.Visible = True Then Exit Sub
    
'---��������� �������� �������
    CommOptForm.ShowForm
Tail:
    '---������� �� ��������� ���������
End Sub


