Attribute VB_Name = "ToolBars"

Sub AddTBImagination()
'��������� ���������� ������ ���������� "�����������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    'Const DocPath = ThisDocument.Path
    'Dim DocPath As String
    
'---��������� ���� �� ��� ������ ���������� "�����������"------------------------------
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = "�����������" Then Exit Sub
    Next i

'---������� ������ ���������� "�����������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "�����������"
        .Visible = True
    End With

End Sub

Sub RemoveTBImagination()
'��������� ���������� ������ ���������� "�����������"-------------------------------
    Application.CommandBars("�����������").Delete
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error GoTo EX
    
    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "�������� � ��������� ����"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "��������� ����"
        .tag = "CalcArea"
        .TooltipText = "�������� � ��������� ����"
'        .Picture = LoadPicture(DocPath & "Bitmaps\Fire1.bmp")
'        .Mask = LoadPicture(DocPath & "Bitmaps\Fire2.bmp")
        .FaceID = 150
        .BeginGroup = True
    End With
'---������ "�������� � ���� �������"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "�������"
        .tag = "FireAreae"
        .TooltipText = "�������� � ���� �������"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fire1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fire2.bmp")
        .BeginGroup = False
    End With
'---������ "�������� � ���� �������"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "�����"
        .tag = "FireStorm"
        .TooltipText = "�������� � �������� �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\Storm1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Storm2.bmp")
    End With
'---������ "�������� � ����������"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "����������"
        .tag = "Fog"
        .TooltipText = "�������� � ����������� ����"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fog1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fog2.bmp")
    End With
'---������ "�������� � ���� ���������"-------------------------------------------------
'    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Bar.Controls.Add(Type:=msoControlButton)
        .Caption = "���������"
        .tag = "Rush"
        .TooltipText = "�������� � ���� ���������"
        .Picture = LoadPicture(DocPath & "Bitmaps\Rush1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Rush2.bmp")
    End With
    
    
    Set Button = Nothing
    Set Bar = Nothing

Exit Sub
EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub


Sub DeleteButtons()
'---��������� �������� ������ "�������" �� ������ ���������� "�����������"--------------
'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String
    
    On Error Resume Next

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "��������� ����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("��������� ����")
    Button.Delete
'---�������� ������ "�������" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("�������")
    Button.Delete
'---�������� ������ "�����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("�����")
    Button.Delete
'---�������� ������ "����������" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("����������")
    Button.Delete
'---�������� ������ "���������" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("���������")
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
