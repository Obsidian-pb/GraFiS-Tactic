VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut3 As Office.CommandBarButton
Attribute ComBut3.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut3 = Application.CommandBars("�����������").Controls("����������")
End Sub

Private Sub Class_Terminate()
    Set ComBut3 = Nothing
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

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
            '---�������� ������ � ���� ����������
                MorphToFog
            End If
        Else
'            MsgBox "�� ������� �� ����� ������! ����� ������� ����� ��������������� ���������!"
            PS_CheckButtons Ctrl
        End If
    End If
    

    
Exit Sub
EX:
    SaveLog Err, "ComBut3_Click"
End Sub

Public Sub MorphToFog()
'�������� ������ � ���� ����������

    On Error GoTo EX
'---��������� ������� �� ����������� ��������� ������
'    If Not CheckSquareShape Then
'    '---�������� ������
'        PS_CheckButtons ComBut3
'        Exit Sub
'    End If
    
'---��������� ��������� ��������� ��������� ������ � ���� ����������
    ImportFogInformation
    
Exit Sub
EX:
    SaveLog Err, "MorphToFog"
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("����������")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Fog1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fog2.bmp")
    End With
    
Set cb = Nothing
End Sub
