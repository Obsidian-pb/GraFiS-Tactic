Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTB_f()
'��������� ���������� ������ ���������� "�������"-------------------------------
Dim i As Integer
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---��������� ���� �� ��� ������ ���������� "�������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).name = "�������" Then Exit Sub
    Next i

'---������� ������ ���������� "�������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .name = "�������"
        .visible = True
    End With

End Sub

Sub RemoveTB_f()
'��������� ���������� ������ ���������� "�������"-------------------------------
    On Error Resume Next
    Application.CommandBars("�������").Delete
End Sub

Sub AddButtons_f()
'��������� ���������� ����� ������ �� ������ ���������� "�������"--------------
    
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("�������")
    
'---��������� ������ �� ������ ���������� "�������"--------------------------------
'---������ "��������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������"
        .Tag = "Refresh all formulas"
        .TooltipText = "�������� ��� ������� �� �����"
        .FaceID = 37
    End With
'---������ "�������� ���"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "�������� ���"
        .Tag = "Show all formulas"
        .TooltipText = "�������� ��� ���������� � ����� ����"
        .FaceID = 139
    End With
    
    Set btns = New c_Buttons
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
ex:
    Set Button = Nothing
    Set Bar = Nothing
'    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.name
    SaveLog Err, "AddButtons_f", "������ �� ������ �������"
End Sub


Sub DeleteButtons_f()
'---��������� �������� ������ �� ������ ���������� "�������"--------------
    On Error GoTo ex
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("�������")
'---�������� ������ "��������" �� ������ ���������� "�������"------------------------
    Set Button = Bar.Controls("��������")
    Button.Delete
'---�������� ������ "��������" �� ������ ���������� "�������"------------------------
    Set Button = Bar.Controls("�������� ���")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
ex:
'������� �� ���������
    Set Button = Nothing
    Set Bar = Nothing
    SaveLog Err, "AddButtons_f", "������ �� ������ �������"
End Sub
