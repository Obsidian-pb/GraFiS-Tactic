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
Public WithEvents ComButClearComInf As Office.CommandBarButton       '��������
Attribute ComButClearComInf.VB_VarHelpID = -1

Public WithEvents ComButMark As Office.CommandBarButton              '������
Attribute ComButMark.VB_VarHelpID = -1

Public WithEvents ComButDescView As Office.CommandBarButton          '�������� �������� ��
Attribute ComButDescView.VB_VarHelpID = -1
Public WithEvents ComButTechView As Office.CommandBarButton          '�������� ������ �������
Attribute ComButTechView.VB_VarHelpID = -1
Public WithEvents ComButPersonnel As Office.CommandBarButton         '�������� �������� ������� ������� � ������ �������
Attribute ComButPersonnel.VB_VarHelpID = -1
Public WithEvents ComButNozzlesView As Office.CommandBarButton       '�������� ������ �������
Attribute ComButNozzlesView.VB_VarHelpID = -1
Public WithEvents ComButGDZSView As Office.CommandBarButton          '�������� ������ ����
Attribute ComButGDZSView.VB_VarHelpID = -1
Public WithEvents ComButTimelineView As Office.CommandBarButton      '�������� ��������
Attribute ComButTimelineView.VB_VarHelpID = -1
Public WithEvents ComButDutyListView As Office.CommandBarButton      '�������� ��������
Attribute ComButDutyListView.VB_VarHelpID = -1
Public WithEvents ComButStatistsView As Office.CommandBarButton      '�������� ������ ���������
Attribute ComButStatistsView.VB_VarHelpID = -1
Public WithEvents ComButPosredniksView As Office.CommandBarButton    '�������� ������ �����������
Attribute ComButPosredniksView.VB_VarHelpID = -1
Public WithEvents ComButExplicationView As Office.CommandBarButton   '�������� �����������
Attribute ComButExplicationView.VB_VarHelpID = -1
Public WithEvents NodesListComBut As Office.CommandBarButton         '������ ����� ���� ���������
Attribute NodesListComBut.VB_VarHelpID = -1

Public WithEvents ComButDescExport As Office.CommandBarButton        '������� �������� ��
Attribute ComButDescExport.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComButCommand = Application.CommandBars("���").Controls("�������")
    Set ComButInfo = Application.CommandBars("���").Controls("����������")
    Set ComButClearComInf = Application.CommandBars("���").Controls("��������")
    Set ComButMark = Application.CommandBars("���").Controls("������")
    
    Set ComButDescView = Application.CommandBars("���").Controls("�������� �������� ��")
    Set ComButTechView = Application.CommandBars("���").Controls("������ �������")
    Set ComButPersonnel = Application.CommandBars("���").Controls("������ ������� �������")
    Set ComButNozzlesView = Application.CommandBars("���").Controls("������ �������")
    Set ComButGDZSView = Application.CommandBars("���").Controls("������ ����")
    Set ComButTimelineView = Application.CommandBars("���").Controls("��������")
    Set ComButDutyListView = Application.CommandBars("���").Controls("����������� ����")
    
    Set ComButStatistsView = Application.CommandBars("���").Controls("������ ���������")
    Set ComButPosredniksView = Application.CommandBars("���").Controls("������ �����������")
    
    Set ComButExplicationView = Application.CommandBars("���").Controls("�����������")
    
    Set ComButDescExport = Application.CommandBars("���").Controls("������� �������� ��")
    
    On Error Resume Next
    Set NodesListComBut = Application.CommandBars("���").Controls("���� ���������")
End Sub

Private Sub Class_Terminate()
    Set ComButCommand = Nothing
    Set ComButInfo = Nothing
    Set ComButClearComInf = Nothing
    Set ComButMark = Nothing
    
    Set ComButDescView = Nothing
    Set ComButTechView = Nothing
    Set ComButPersonnel = Nothing
    Set ComButNozzlesView = Nothing
    Set ComButGDZSView = Nothing
    Set ComButTimelineView = Nothing
    Set ComButDutyListView = Nothing
    
    Set ComButStatistsView = Nothing
    Set ComButPosredniksView = Nothing
    
    Set ComButExplicationView = Nothing
    
    Set ComButDescExport = Nothing
    
    On Error Resume Next
    Set NodesListComBut = Nothing
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

Private Sub ComButClearComInf_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ClearComInf
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



'-------------------------������ �������--------------------------------------------------
Private Sub ComButDescView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    DescriptionViewInList
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButDutyListView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowDutyList
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButPosredniksView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowPosredniks
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComButTechView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowUnits
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButPersonnel_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowPersonnel
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButNozzlesView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowNozzles
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButGDZSView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowGDZS
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButTimelineView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowTimeLine
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButStatistsView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowStatists
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub ComButExplicationView_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowExplication
Tail:
    '---������� �� ��������� ���������
End Sub
Private Sub NodesListComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
    
'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShowEvacNodes
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
