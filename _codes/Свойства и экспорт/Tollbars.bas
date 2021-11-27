Attribute VB_Name = "Tollbars"
Option Explicit

Private btns As c_Buttons


Sub AddTB()
'��������� ���������� ������ ���������� "�������"-------------------------------
Dim i As Integer
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
'---��������� ���� �� ��� ������ ���������� "�������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "�������" Then Exit Sub
    Next i

'---������� ������ ���������� "�������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "�������"
        .Visible = True
    End With
'---��������� ������
    AddButtons
    
End Sub

Sub RemoveTB()
'��������� ���������� ������ ���������� "�������"-------------------------------
    On Error Resume Next
    DeleteButtons
    Application.CommandBars("�������").Delete
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "�������"--------------
    
'---��������� ���������� � ����������--------------------------------------------------
Dim Bar As CommandBar
Dim DocPath As String
    
    On Error GoTo ex
    
    Set Bar = Application.CommandBars("�������")
    
'---��������� ������ �� ������ ���������� "�������"--------------------------------
'---������ "���������"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "���������"
        .Tag = "Report"
        .TooltipText = "������������ ��������� � ������"
        .FaceID = 626
    End With
'---������ "���"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "���"
        .Tag = "KBD"
        .TooltipText = "������������ �������� ������ ��������"
        .FaceID = 626
    End With
    
    Set btns = New c_Buttons
    
    Set Bar = Nothing

Exit Sub
ex:
    Set Bar = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddButtons", "�������"
End Sub


Sub DeleteButtons()
'---��������� �������� ������ "�������" �� ������ ���������� "�������"--------------
    On Error GoTo ex
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar

    Set Bar = Application.CommandBars("�������")
'---�������� ������ "�������" �� ������ ���������� "�������"------------------------
    Bar.Controls("���������").Delete

Set btns = Nothing

Set Bar = Nothing

Exit Sub
ex:
'������� �� ���������
End Sub

