VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut4 As Office.CommandBarButton
Attribute ComBut4.VB_VarHelpID = -1


Private Sub Class_Initialize()
'    MsgBox "123"
    Set ComBut4 = Application.CommandBars("�����������").Controls("���������")
End Sub

Private Sub Class_Terminate()
    Set ComBut4 = Nothing
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    On Error GoTo EX
    
'---��������� �������� �� ������ - ���� �������� - ���������
    If Ctrl.state = msoButtonDown Then
        Ctrl.state = msoButtonUp
        Exit Sub
    End If

    If Ctrl.state = msoButtonUp Then
        If IsSelectedOneShape(False) Then
        '---���� ������� ���� ���� ������ - �������� �� ��������
            If IsHavingUserSection(True) And IsSquare(True) Then
            '---�������� � ���� ���������
                MorphToRush
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
EX:
    SaveLog Err, "ComBut4_Click"
End Sub

Public Sub MorphToRush()

    On Error GoTo EX
'---��������� ������� �� ����������� ��������� ������
'    If Not CheckRushShape Then
'    '---�������� ������
'        PS_CheckButtons ComBut4
'        Exit Sub
'    End If

'---��������� ��������� ��������� ��������� ������ � ���� ���������
    ImportRushInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToRush"
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("���������")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Rush1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Rush2.bmp")
    End With
    
Set cb = Nothing
End Sub
