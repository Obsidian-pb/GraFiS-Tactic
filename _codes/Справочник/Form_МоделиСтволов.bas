VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_�������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_BeforeUpdate(Cancel As Integer)
'On Error Resume Next
'    TB_LastChangedTime.Value = Now()
'Dim rst As Recordset

'    Set rst = Me.Recordset
'    rst.Edit
'        rst.Fields("��������").Value = Now()
'    rst.Update

End Sub

Private Sub Form_Current()
'��������� �������� ����� ��������
'---��������� �������� ����������
    �_��������������������.Value = False
    s_ControlsBlockChange
    
'---��������������� ������� ������� ��������� ����������
'    Me.�_���������������������.Enabled = False

'---��������� ������ ��������� � ����� �������
    Me.�_�����������������.Requery
    Me.�_�����������������.Value = Me.�_�����������������.ItemData(0)
    Me.��_������������.Requery
    
'---����������� ���� ��������� �������� �� ������ �_�����������������
    Me.�_��������� = Me.�_�����������������.Column(3)
    
End Sub

Private Sub Form_Load()
    Me.���_�����������.Enabled = False
    Me.���_�����������.Value = 1
    Me.�_������������������.Value = False
    Me.���_�������.Value = " "
    Me.���_�������.Requery
End Sub

Private Sub �_��������������������_AfterUpdate()
    s_ControlsBlockChange
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub �_���������������������_Click()
    DoCmd.OpenForm "���������������", acNormal
End Sub

Private Sub �_���������_AfterUpdate()
'��������� ���������� �������� ���� ��������� � ������� ���������������
'---��������� ������ �� ��� �������� ������ ������� � ���� ��� ������ �������������� � �������
    If Me.�_�����������������.ListCount = 0 Then
        MsgBox "������� ���������� ������� ���� ���� ������� ������!", vbInformation
        �_��������� = ""
        Exit Sub
    End If
    
'---��������� ����������
    Dim rst As DAO.Recordset, dbs As DAO.Database
    
'---����������� ���������� �������
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("���������������")
    
'---���� ������ � ������� �������� ������� ��������������� ���� ���������� � ������ ������� ������ �_�����������������_
'    � ����������� �� �������� ��������� ������������� � ���� �_���������
    With rst
        .FindFirst ("����������������� =" & Me.�_�����������������.Column(0))
        .Edit
        ![���������] = �_���������.Value
        .Update
    End With
    
    rst.Close
    dbs.Close
    
    Me.�_�����������������.Requery
End Sub

Private Sub ���_�������_AfterUpdate()
    ' ����� ������, ��������������� ����� �������� ����������.
    Dim rs As Object

'    On Error GoTo EX

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[���������������] = " & str(Nz(Me.���_�������.Value, 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
'EX:
End Sub

Private Sub ���_�����������_AfterUpdate()
    Me.���_�������.Requery
    Me.Requery
End Sub


Private Sub �_�����������������_AfterUpdate()
    Me.��_������������.Requery
    Me.�_��������� = Me.�_�����������������.Column(3)
End Sub

Private Sub s_ControlsBlockChange()
'��������� ����������-������������� ��������� ���������� � ����� �������������
'---��� ��������� �����
    Me.�_������������.Locked = Not �_��������������������.Value
    Me.���_�������������.Locked = Not �_��������������������.Value
    Me.�_��������������.Locked = Not �_��������������������.Value
    Me.�_���������.Locked = Not �_��������������������.Value
    Me.�_���������������������.Enabled = �_��������������������.Value
    Me.�_������_WF.Locked = Not �_��������������������.Value
    Me.�_��������_��������.Locked = Not �_��������������������.Value
    
'---��� ����������� ����
    ��_������������.Locked = Not �_��������������������.Value
    Me.��_������������.Controls(2).Locked = Not �_��������������������.Value
    
'---��� ������ ��������� �������: ����������� �������� ���������
    If �_��������������������.Value = True Then
        Me.�_�����������������.ListItemsEditForm = "���������������"
    Else
        Me.�_�����������������.ListItemsEditForm = ""
    End If
    
'---��� ��_�����������
    Me.��_������������.Controls("��� �����").Locked = Not �_��������������������.Value
End Sub


Private Sub �_������������������_AfterUpdate()
    Me.���_�����������.Enabled = �_������������������
    Me.���_�������.Requery
    Me.Requery
End Sub


