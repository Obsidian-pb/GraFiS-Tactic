VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����� ��� ��������� ������� �� ������

Public WithEvents ComBut As Office.CommandBarButton   'Master check
Attribute ComBut.VB_VarHelpID = -1



Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("�����������").Controls("������ ��������")


End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    show_m_chek_form
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

