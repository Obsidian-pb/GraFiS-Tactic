Attribute VB_Name = "ToolBars"

Sub AddTBImagination()
'��������� ���������� ������ ���������� "�����������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    'Const DocPath = ThisDocument.Path
    Dim DocPath As String
    
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

Sub RemoveTBImagination()
'��������� ���������� ������ ���������� "�����������"-------------------------------
    On Error Resume Next
    
    Application.CommandBars("�����������").Delete
End Sub



Sub AddTBNRSCalc()
'��������� ���������� ������ ���������� "�������-�������� �������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
'---��������� ���� �� ��� ������ ���������� "�������-�������� �������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "�������-�������� �������" Then Exit Sub
    Next i

'---������� ������ ���������� "�����������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "�������-�������� �������"
        .Visible = True
    End With

End Sub

Sub RemoveTBNRSCalc()
'��������� ���������� ������ ���������� "�������-�������� �������"-------------------------------
    On Error Resume Next
    
    Application.CommandBars("�������-�������� �������").Delete
End Sub

'--------------------------------------������ ������� �����-------------------------
Sub AddButtonLine()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path

'---��������� ���� �� ��� �� ������ ���������� "�����������" ������ "�����"------------------------------
'    For i = 1 To Application.CommandBars("�����������").Controls.Count
''    Application.CommandBars("�����������").Controls
'        If Application         'CommandBars("�����������").Controls(i).Name = "�����" Then

'    Next i
    
'    If Not Application.CommandBars("�����������").Controls("�����") = Nothing Then
'        Set Bar = Nothing
'        Exit Sub
'    End If
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "�������� � �������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "�����"
        .Tag = "Hose"
        '.OnAction = "Application.Documents('����� ���.vss').ExecuteLine ('ProvExchange')"
        .TooltipText = "�������� � ������� �������� �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\Hose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Hose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub


Sub DeleteButtonLine()
'---��������� �������� ������ "�����" � ������ ���������� "�����������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "�����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("�����")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------������ ������������� �����-------------------------
Sub AddButtonMLine()
'��������� ���������� ����� ������ "������������� �����" �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "�������� � ������������� �������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������������� �����"
        .Tag = "MHose"
        .TooltipText = "�������� � ������������� �������� �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\MHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MHose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonMLine()
'---��������� �������� ������ "������������� �����" � ������ ���������� "�����������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "�����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("������������� �����")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub


'--------------------------------------������ ����������� �����-------------------------
Sub AddButtonVHose()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "�������� �� ����������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����������� �����"
        .Tag = "VHose"
        .TooltipText = "�������� �� ����������� �������� �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\VHose1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\VHose2.bmp")
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonVHose()
'��������� �������� ������ "����������� �����" � ������ ���������� "�����������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "�����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("����������� �����")
    Button.Delete

    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------������ ������ ���-------------------------
Sub AddButtonNormalize()
'��������� ���������� ����� ������ �� ������ ���������� "�������-�������� �������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
    DocPath = ThisDocument.path
'---��������� ������ �� ������ ���������� "�������-�������� �������"--------------------------------
'---������ "������������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������ ���"
        .Tag = "Calculate NRS"
        .TooltipText = "���������� ���"
        .FaceID = 807
        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNormalize()
'��������� �������� ������ "������������" � ������ ���������� "�������-�������� �������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
'---�������� ������ "������������" �� ������ ���������� "�������-�������� �������"------------------------
    Set Button = Bar.Controls("������ ���")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------������ ��������� ������� ���-------------------------
Sub AddButtonNRSSettings()
'��������� ���������� ����� ������ �� ������ ���������� "�������-�������� �������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
    DocPath = ThisDocument.path
'---��������� ������ �� ������ ���������� "�������-�������� �������"--------------------------------
'---������ "��������� ������� ���"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������� ������� ���"
        .Tag = "NRS Settings"
        .TooltipText = "��������� ������� ���"
        .FaceID = 642
'        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNRSSettings()
'��������� �������� ������ "��������� ������� ���" � ������ ���������� "�������-�������� �������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
'---�������� ������ "������������" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("��������� ������� ���")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub

'--------------------------------------������ ����� ������� ���-------------------------
Sub AddButtonNRSReport()
'��������� ���������� ����� ������ �� ������ ���������� "�������-�������� �������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
    DocPath = ThisDocument.path
'---��������� ������ �� ������ ���������� "�������-�������� �������"--------------------------------
'---������ "��������� ������� ���"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����� ������� ���"
        .Tag = "NRS Report"
        .TooltipText = "�������� ����� ������� ���"
        .FaceID = 230
'        .BeginGroup = True
    End With
    Set Button = Nothing

Set Bar = Nothing
End Sub

Sub DeleteButtonNRSReport()
'��������� �������� ������ "����� ������� ���" � ������ ���������� "�������-�������� �������"--------------
    
    On Error Resume Next
    
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�������-�������� �������")
'---�������� ������ "����� ������� ���" �� ������ ���������� "�������-�������� �������"------------------------
    Set Button = Bar.Controls("����� ������� ���")
    Button.Delete
    
Set Button = Nothing
Set Bar = Nothing
End Sub



'-----------------------------����������� ������ � ��������-------------------------------------
Public Sub PS_CheckButtons(ByRef a_MainBtn As Office.CommandBarButton)
'��������� ��������� ���������� ������ ��������� ������ (� ��� ������, ���� �� ������� �� ����� ������)
Dim v_Cntrl As CommandBarControl
    
    If Application.ActiveWindow.Selection.Count >= 1 Then Exit Sub
    
    For Each v_Cntrl In Application.CommandBars("�����������").Controls
        If v_Cntrl.Caption = a_MainBtn.Caption Then
            v_Cntrl.State = Not a_MainBtn.State 'msoButtonDown
        Else
            v_Cntrl.State = False 'msoButtonUp
        End If
    Next v_Cntrl
End Sub
