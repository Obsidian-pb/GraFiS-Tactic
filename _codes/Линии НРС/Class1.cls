VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton   '������� �����
Attribute ComBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  '����������� �����
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  '������������� �����
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  '������ ���
Attribute ComBut4.VB_VarHelpID = -1
Public WithEvents ComBut5 As Office.CommandBarButton  '��������� ������� ���
Attribute ComBut5.VB_VarHelpID = -1
Public WithEvents ComBut6 As Office.CommandBarButton  '����� ������� ���
Attribute ComBut6.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("�����������").Controls("�����")
    Set ComBut2 = Application.CommandBars("�����������").Controls("����������� �����")
    Set ComBut3 = Application.CommandBars("�����������").Controls("������������� �����")
    Set ComBut4 = Application.CommandBars("�������-�������� �������").Controls("������ ���")
    Set ComBut5 = Application.CommandBars("�������-�������� �������").Controls("��������� ������� ���")
    Set ComBut6 = Application.CommandBars("�������-�������� �������").Controls("����� ������� ���")
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
    Set ComBut2 = Nothing
    Set ComBut3 = Nothing
    Set ComBut4 = Nothing
    Set ComBut5 = Nothing
    Set ComBut6 = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "�������� � �������� �����"
    
'    On Error GoTo EX

'---���������
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub

'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---�������� ������ � ������ �����
                MakeHoseLine 51, 0
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    
'            '---�������� ������ � ������ ���� �������
'                MorphToLake

Exit Sub
ex:
    SaveLog Err, "ComBut_Click"
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ����� �� ������ "����������� �����"

    On Error GoTo ex

'---���������
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub
    
'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---�������� ������ � ������ ����������� �����
                MakeVHoseLine
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
ex:
    SaveLog Err, "ComBut2_Click"
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ����� �� ������ "������������� �����"

    On Error GoTo ex
    
'---���������
'    If ClickAndOnSameButton(Ctrl.Caption) Then Exit Sub
    
'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---�������� ������ � ������ ������������� �����
                MakeHoseLine 77, 1
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If

Exit Sub
ex:
    SaveLog Err, "ComBut3_Click"
End Sub



'--------------------------------���������� ������� ������---------------------
Private Function GetSelectedButtonCaption() As String
Dim v_Ctrl As CommandBarControl

    On Error GoTo ex
'---��������� ����� ������ ������ � � ����������� �� ����� ��������� ��������
    For Each v_Ctrl In Application.CommandBars("�����������").Controls
        If v_Ctrl.State = msoButtonDown Then
            GetSelectedButtonCaption = v_Ctrl.Caption
            Exit Function
        End If
    Next v_Ctrl
GetSelectedButtonCaption = ""

Exit Function
ex:
End Function



Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'������ ���
    
    'NormalizeNRS
Dim NRSDemon As c_NRSDemon
Dim shp As Visio.Shape

    On Error GoTo ex
    
    If Not Application.ActiveWindow.Selection.Count = 1 Then Exit Sub
    
    Set shp = Application.ActiveWindow.Selection(1)
    
    Set NRSDemon = New c_NRSDemon

    NRSDemon.roundAccuracy = GetSetting("GraFiS", "GFS_NRS", "RoundAccuracy", 4)
    NRSDemon.checkAccuracy = GetSetting("GraFiS", "GFS_NRS", "CheckAccuracy", 2)
    NRSDemon.outAccuracy = GetSetting("GraFiS", "GFS_NRS", "OutAccuracy", 2)
    NRSDemon.maxIterations = GetSetting("GraFiS", "GFS_NRS", "MaxIterations", 100)
    NRSDemon.approvedHout = GetSetting("GraFiS", "GFS_NRS", "ApprovedHout", 3)


    '������ ������
    NRSDemon.BuildNRSModel shp
    '������������ ������
    NRSDemon.CalculateNRSModel
    
    Set NRSDemon = Nothing
Exit Sub
ex:
    Set NRSDemon = Nothing
End Sub

Private Sub ComBut5_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'��������� ������� ���
    
    frmNRSSettings.Show
    
End Sub

Private Sub ComBut6_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����� ������� ���
    
    frmNRSDescription.Show
    
End Sub
