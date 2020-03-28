Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons

Sub AddTB_SpecFunc()
'��������� ���������� ������ ���������� "�����������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---��������� ���� �� ��� ������ ���������� "�����������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "�����������" Then Exit Sub
    Next i

'---������� ������ ���������� "�����������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "�����������"
        .Visible = True
    End With

End Sub

Sub RemoveTB_SpecFunc()
'��������� ���������� ������ ���������� "�����������"-------------------------------
    On Error GoTo EX
    Application.CommandBars("�����������").Delete
    Set btns = Nothing
EX:
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------
    On Error GoTo EX
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "������ ��������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������ ��������"
        .Tag = "show_m_chek_form"
        .TooltipText = "������ ��������"
        .FaceID = 1820
        .beginGroup = True
    End With
'---������ "����������� ������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����������� ������"
        .Tag = "show_m_tactic_form"
        .TooltipText = "����������� ������"
        .FaceID = 1090
        .beginGroup = False
    End With
'---������ "���������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "���������"
        .Tag = "calculationSettings"
        .TooltipText = "���������"
        .FaceID = 642
        .beginGroup = False
    End With
    
    
    Set Button = Nothing
    
'---���������� ����� ������������� ������
    Set btns = New c_Buttons
    
Set Bar = Nothing
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "AddButtons"
End Sub

Sub DeleteButtons()
'---��������� �������� ������ � ������ ���������� "�����������"--------------
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "������ ��������" � ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("������ ��������")
    Button.Delete
'---�������� ������ "����������� ������" � ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("����������� ������")
    Button.Delete
'---�������� ������ "���������" � ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("���������")
    Button.Delete
    
'---���������� ����� ������������� ������
    Set btns = Nothing
    
Set Button = Nothing
Set Bar = Nothing
End Sub
