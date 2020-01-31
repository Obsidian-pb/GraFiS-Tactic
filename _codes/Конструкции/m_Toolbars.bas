Attribute VB_Name = "m_Toolbars"
Sub AddTB_Constructions()
'��������� ���������� ������ ���������� "�����������"-------------------------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim i As Integer
    
'---��������� ���� �� ��� ������ ���������� "�����������"------------------------------
    For i = 1 To Application.CommandBars.count
        If Application.CommandBars(i).Name = "�����������" Then Exit Sub
    Next i

'---������� ������ ���������� "�����������"--------------------------------------------
    Set Bar = Application.CommandBars.Add(Position:=msoBarRight, Temporary:=True)
    With Bar
        .Name = "�����������"
        .Visible = True
    End With
'    AddButtons
End Sub

Sub RemoveTB_Constructions()
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
'---������ "�����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .caption = "�����"
        .Tag = "WallsMask"
        .TooltipText = "�������� ����� ����"
        .Picture = LoadPicture(DocPath & "Bitmaps\WallMask1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallMask2.bmp")
    End With
'---������ "��������� ����"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .caption = "��������� ����"
        .Tag = "WallDrawer"
        .TooltipText = "��������� ����"
        .Picture = LoadPicture(DocPath & "Bitmaps\WallDrawTool1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallDrawTool2.bmp")
    End With

    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "AddButtons"
End Sub
