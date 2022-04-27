Attribute VB_Name = "m_Toolbars"
Option Explicit



Public Sub AddTB_Evacuation()
'��������� ���������� ������ ���������� "���������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim i As Integer
    
'---��������� ���� �� ��� ������ ���������� "���������"------------------------------
    For i = 1 To Application.CommandBars.count
        If Application.CommandBars(i).Name = "���������" Then Exit Sub
    Next i

'---������� ������ ���������� "���������"--------------------------------------------
    With Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
        .Name = "���������"
        .Visible = True
    End With
    
    AddButtons
End Sub

Public Sub RemoveTB_Evacuation()
'��������� ���������� ������ ���������� "���������"-------------------------------
    On Error GoTo EX
    Application.CommandBars("���������").Delete

EX:
End Sub

Public Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "���������"--------------
    On Error GoTo EX
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("���������")
    DocPath = ThisDocument.path
    
'---��������� ������ �� ������ ���������� "���������"--------------------------------
'---������ "������� ��� ������ �����"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "������� ���"
        .Tag = "SelectAllGraphShapes"
        .TooltipText = "������� ��� ������ �����"
        .FaceID = 1446
    End With
'---������ "�������������� ��� ������ �����"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "��������������"
        .Tag = "Renum"
        .TooltipText = "�������������� ��� ������ �����"
        .FaceID = 1116
    End With

'---������ "���������� ����� ���������"-------------------------------------------------
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "����������"
        .Tag = "Calculate"
        .TooltipText = "���������� ����� ���������"
        .FaceID = 283
        .BeginGroup = True
    End With
    

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub

