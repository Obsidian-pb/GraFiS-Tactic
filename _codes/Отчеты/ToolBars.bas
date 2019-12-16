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
'    AddButtons
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
        .Caption = "������ ��������"
        .Tag = "show_m_chek_form"
        .TooltipText = "��������� ������������ �����"
        .Picture = LoadPicture(DocPath & "Bitmaps\MasterCheck.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\MasterCheck2.bmp")
    End With

    
    
    Set Button = Nothing

Set Bar = Nothing
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "AddButtons"
End Sub

