Attribute VB_Name = "Toolbars"
Sub AddTBImagination()
'��������� ���������� ������ ���������� "�����������"-------------------------------
Dim i As Integer

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    
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
    Application.CommandBars("�����������").Delete
End Sub

Sub AddButtons()
'��������� ���������� ����� ������ �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������--------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
    DocPath = ThisDocument.path
    
'---��������� ������ �� ������ ���������� "�����������"--------------------------------
'---������ "�������� � �������� ������������"-------------------------------------------------
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "������������ ������������"
        .Tag = "NaturalWater"
        .TooltipText = "�������� � ������������ ������������"
        .Picture = LoadPicture(DocPath & "Bitmaps\Lake1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Lake2.bmp")
        .BeginGroup = True
    End With
    
    Set Button = Nothing

Set Bar = Nothing
End Sub


Sub DeleteButtons()
'---��������� �������� ������ "�������" �� ������ ���������� "�����������"--------------

'---��������� ���������� � ����������-------------------------------------------------
    Dim Bar As CommandBar, Button As CommandBarButton
    Dim DocPath As String

    Set Bar = Application.CommandBars("�����������")
'---�������� ������ "�����" �� ������ ���������� "�����������"------------------------
    Set Button = Bar.Controls("������������ ������������")
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
            v_Cntrl.state = Not a_MainBtn.state 'msoButtonDown
        Else
            v_Cntrl.state = False 'msoButtonUp
        End If
    Next v_Cntrl
End Sub
