VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFireArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton
Attribute ComBut.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("�����������").Controls("�������")
'MsgBox ComBut.Caption
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

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
            '---�������� ������ � ������ ���� �������
                MorphToFireArea
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    
Exit Sub
EX:
    SaveLog Err, "ComBut_Click"
End Sub

Public Sub MorphToFireArea()

    On Error GoTo EX

'---��������� ������� �� ����������� ��������� ������
'    If Not CheckSquareShape Then
'    '---�������� ������
'        PS_CheckButtons ComBut 'Ctrl
'        Exit Sub
'    End If

'---��������� ��������� ��������� ��������� ������ � ������� ������
    ImportAreaInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToFireArea"
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("�������")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Fire1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fire2.bmp")
    End With
    
Set cb = Nothing
End Sub
