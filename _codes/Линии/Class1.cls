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

Public WithEvents comBut As Office.CommandBarButton   '������� �����
Attribute comBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  '����������� �����
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  '������������� �����
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  '������� �����
Attribute ComBut4.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set comBut = Application.CommandBars("�����������").Controls("�����")
    Set ComBut2 = Application.CommandBars("�����������").Controls("����������� �����")
    Set ComBut3 = Application.CommandBars("�����������").Controls("������������� �����")
    Set ComBut4 = Application.CommandBars("�����������").Controls("�������")
End Sub

Private Sub Class_Terminate()
    Set comBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "�������� � �������� �����"
    
    On Error GoTo EX

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
                MakeHoseLine Application.ActiveWindow.Selection(1), 51, 0
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
    SaveLog Err, "ComBut_Click"
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ����� �� ������ "����������� �����"

    On Error GoTo EX
    
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
EX:
    SaveLog Err, "ComBut2_Click"
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ����� �� ������ "������������� �����"

    On Error GoTo EX
        
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
                MakeHoseLine Application.ActiveWindow.Selection(1), 77, 1
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If

Exit Sub
EX:
    SaveLog Err, "ComBut3_Click"
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ����� �� ������ "�������"
    
    On Error GoTo EX
    
'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.State = msoButtonDown Then
        Ctrl.State = msoButtonUp
        Exit Sub
    End If

    If Ctrl.State = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If Not IsHavingUserSection(True) And Not IsSquare(True) Then
            '---�������� ������ � ������ ��������
                MakeHoseDrop Application.ActiveWindow.Selection(1), 51, 1
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
Exit Sub
EX:
    SaveLog Err, "ComBut4_Click"
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("�����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Hose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Hose2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("����������� �����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\VHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\VHose2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("������������� �����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\MHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MHose2.bmp")
    End With
    
    Set cb = Nothing
End Sub



'--------------------------------���������� ������� ������---------------------
Private Function GetSelectedButtonCaption() As String
Dim v_Ctrl As CommandBarControl

    On Error GoTo EX
'---��������� ����� ������ ������ � � ����������� �� ����� ��������� ��������
    For Each v_Ctrl In Application.CommandBars("�����������").Controls
        If v_Ctrl.State = msoButtonDown Then
            GetSelectedButtonCaption = v_Ctrl.Caption
            Exit Function
        End If
    Next v_Ctrl
GetSelectedButtonCaption = ""

Exit Function
EX:
End Function



