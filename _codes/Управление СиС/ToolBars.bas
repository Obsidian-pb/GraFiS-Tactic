Attribute VB_Name = "ToolBars"
Option Explicit

Private btns As c_Buttons


Sub AddTB()
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
        .visible = True
    End With

'---��������� ������ �� ������ ����������
    AddButtons
End Sub

Sub RemoveTB()
'��������� ���������� ������ ���������� "���"-------------------------------
    On Error Resume Next
    Application.CommandBars("���").Delete
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "���"--------------
    
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("���")
    
'---��������� ������ �� ������ ���������� "���"--------------------------------
'---������ "�������"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "�������"
        .Tag = "Command"
        .TooltipText = "������� ����������� �������"
        .FaceID = 346
    End With
'---������ "����������"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "����������"
        .Tag = "Info"
        .TooltipText = "���������� ��� ������"
        .FaceID = 487 ' 162
        .beginGroup = True
    End With
'---������ "������"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������"
        .Tag = "Mark"
        .TooltipText = "������ ��������� ������ �������� ��� ������� �������"
        .FaceID = 215 ' 162
    End With
    
'---������ �������-------------------------------------------------
    '---"��������"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "�������� �������� ��"
        .Tag = "DescriptionView"
        .TooltipText = "�������� �������� ��"
        .beginGroup = True
        .FaceID = 5
'        .Caption = "��������"
    End With
    '---"�������"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������ �������"
        .Tag = "TechView"
        .TooltipText = "�������� ������ �������"
'        .BeginGroup = True
        .FaceID = 1277
    End With
    '---"������"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������ �������"
        .Tag = "NozzlesView"
        .TooltipText = "�������� ������ �������"
'        .BeginGroup = True
        .FaceID = 2644
    End With
    '---"����"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������ ����"
        .Tag = "GDZSView"
        .TooltipText = "�������� ������ � ����� ����"
'        .BeginGroup = True
        .FaceID = 1253
    End With
    '---"��������"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "��������"
        .Tag = "TimelineView"
        .TooltipText = "�������� �������� ������"
'        .BeginGroup = True
        .FaceID = 11
    End With
    '---"��������"
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������ ���������"
        .Tag = "StatistsView"
        .TooltipText = "�������� �������� � ���������"
'        .BeginGroup = True
        .FaceID = 2141
    End With
    
    
    
'---������ "������� �������� ��"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������� �������� ��"
        .Tag = "DescriptionExport"
        .TooltipText = "������� �������� �� � Word"
        .beginGroup = True
        .FaceID = 582
    End With
    
    Set btns = New c_Buttons
    
    Set Bar = Nothing

Exit Sub
EX:
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
'---�������� ������ "����������" �� ������ ���������� "���"------------------------
    Set Button = Bar.Controls("����������")
    Button.Delete
'---�������� ������ "������" �� ������ ���������� "���"------------------------
    Set Button = Bar.Controls("������")
    Button.Delete
'---�������� ������ "������� �������� ��" �� ������ ���������� "���"------------------------
    Set Button = Bar.Controls("������� �������� ��")
    Button.Delete

Set btns = Nothing

Set Button = Nothing
Set Bar = Nothing

Exit Sub
EX:
'������� �� ���������
End Sub
