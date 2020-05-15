Attribute VB_Name = "ToolBars"

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

EX:
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------
    On Error GoTo EX
'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "������� � JPG"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������� � JPG"
        .Tag = "Export_JPG"
        .TooltipText = "�������������� ��� ����� � JPG"
        .Picture = LoadPicture(DocPath & "Bitmaps\ExportJPG1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\ExportJPG2.bmp")
        .BeginGroup = True
    End With
'---������ "������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������"
        .Tag = "Aspect"
        .TooltipText = "�������� ������"
        .Picture = LoadPicture(DocPath & "Bitmaps\Aspect1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Aspect2.bmp")
    End With
'---������ "��������� ������������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������� ������������"
        .Tag = "Fix"
        .TooltipText = "��������� ������������ ����� �� �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\Fix1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fix2.bmp")
    End With
'---������ "���������� �����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "���������� �����"
        .Tag = "Count"
        .TooltipText = "�������� ���������� ����� � �������"
        .Picture = LoadPicture(DocPath & "Bitmaps\Count1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Count2.bmp")
    End With
'---������ "������ �������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������"
        .Tag = "Timer"
        .TooltipText = "�������� ������ ������������ '������'"
        .FaceID = 2146
        .BeginGroup = True
    End With
    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "AddButtons"
End Sub


Sub DeleteButtons()
'---��������� �������� ������ "������ ��������" �� ������ ���������� "�����������"--------------

    On Error GoTo EX

'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "�����" �� ������ ���������� "������� � JPG"------------------------
    Set Button = Bar.Controls("������� � JPG")
    Button.Delete
'---�������� ������ "�����" �� ������ ���������� "������"------------------------
    Set Button = Bar.Controls("������")
    Button.Delete
'---�������� ������ "�����" �� ������ ���������� "��������� ������������"------------------------
    Set Button = Bar.Controls("��������� ������������")
    Button.Delete
'---�������� ������ "�����" �� ������ ���������� "���������� �����"------------------------
    Set Button = Bar.Controls("���������� �����")
    Button.Delete
'---�������� ������ "�����" �� ������ ���������� "������"------------------------
    Set Button = Bar.Controls("������")
    Button.Delete
    
    
Set Button = Nothing
Set Bar = Nothing
Exit Sub

EX:
    Set Button = Nothing
    Set Bar = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "DeleteButtons"
End Sub


