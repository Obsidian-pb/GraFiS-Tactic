VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton
Attribute ComBut.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("�����������").Controls("������������ ������������")
'MsgBox ComBut.Caption
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

'    On Error GoTo EX
'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If IsHavingUserSection(True) And IsSquare(True) Then
            '---�������� ������ � ������ ���� �������
                MorphToLake
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    
'            '---�������� ������ � ������ ���� �������
'                MorphToLake
    
Exit Sub
EX:
End Sub

Public Sub MorphToLake()

    On Error GoTo EX
'---��������� ������� �� ����������� ��������� ������
'    If Not CheckSquareShape Then
'    '---�������� ������
''        PS_CheckButtons ComBut
'    '---�������� - ����� ������� ��� �����������
'        ComBut.State = msoButtonUp
'        Exit Sub
'    End If
    
'---��������� �������� �� ������ - ���� �������� - ���������
'    If ComBut.State = msoButtonDown Then
'        ComBut.State = msoButtonUp
'        Exit Sub
'    End If

'    If ComBut.State = msoButtonUp Then
'        If IsSelectedOneShape(False) Then
'        '---���� ������� ���� ���� ������ - �������� �� ��������
'            If IsHavingUserSection(True) And IsSquare(True) Then
'            '---�������� ������ � ������ ���� �������
'                ImportOpenWaterInformation
'            End If
'        Else
'            PS_CheckButtons ComBut
'        End If
'    End If

'---��������� ��������� ��������� ��������� ������ � ������������ ������������
    ImportOpenWaterInformation
    
Exit Sub
EX:
End Sub
