Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTBImagination()
'��������� ���������� ������ ���������� "���"-------------------------------
Dim i As Integer
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---��������� ���� �� ��� ������ ���������� "���"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "���" Then Exit Sub
    Next i

'---������� ������ ���������� "���"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "���"
        .Visible = True
    End With

End Sub

Sub RemoveTBImagination()
'��������� ���������� ������ ���������� "���"-------------------------------
    On Error Resume Next
    Application.CommandBars("���").Delete
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "���"--------------
    
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("���")
    
'---��������� ������ �� ������ ���������� "���"--------------------------------
'---������ "�������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "�������"
        .Tag = "Command"
        .TooltipText = "������� ����������� �������"
        .FaceID = 238
    End With
    
    Set btns = New c_Buttons
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddButtons", "�������"
End Sub


Sub DeleteButtons()
'---��������� �������� ������ "�������" �� ������ ���������� "���"--------------
    On Error GoTo EX
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton

    Set Bar = Application.CommandBars("���")
'---�������� ������ "�������" �� ������ ���������� "���"------------------------
    Set Button = Bar.Controls("�������")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
EX:
'������� �� ���������
End Sub
