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
Public WithEvents ComButCommand As Office.CommandBarButton           '�������
Attribute ComButCommand.VB_VarHelpID = -1
Public WithEvents ComButInfo As Office.CommandBarButton              '����������
Attribute ComButInfo.VB_VarHelpID = -1
Public WithEvents ComButMark As Office.CommandBarButton              '������
Attribute ComButMark.VB_VarHelpID = -1

Public WithEvents ComButDescView As Office.CommandBarButton          '�������� �������� ��
Attribute ComButDescView.VB_VarHelpID = -1
Public WithEvents ComButDescExport As Office.CommandBarButton        '������� �������� ��
Attribute ComButDescExport.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCommand = Application.CommandBars("���").Controls("�������")
    Set ComButInfo = Application.CommandBars("���").Controls("����������")
    Set ComButMark = Application.CommandBars("���").Controls("������")
    
    Set ComButDescView = Application.CommandBars("���").Controls("�������� �������� ��")
    Set ComButDescExport = Application.CommandBars("���").Controls("������� �������� ��")
End Sub

Private Sub Class_Terminate()
    Set ComButCommand = Nothing
    Set ComButInfo = Nothing
    Set ComButMark = Nothing
    
    Set ComButDescView = Nothing
    Set ComButDescExport = Nothing
End Sub


Private Sub ComButCommand_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    frm_Command.NewCommand
Tail:
    '---������� �� ��������� ���������
End Sub


Private Sub ComButDescExport_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    DescriptionExportToWord
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButDescView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    DescriptionViewInList
Tail:
    '---������� �� ��������� ���������
End Sub

'-------------------------������ ����������--------------------------------------------------
Private Sub ComButInfo_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    frm_InfoSimple.NewInfo
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButMark_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    frm_InfoSimple.NewInfo 1
Tail:
    '---������� �� ��������� ���������
End Sub
