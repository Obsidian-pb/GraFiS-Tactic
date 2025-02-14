VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Buttons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����� ��� ��������� ������� �� ������
Private WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1

Private WithEvents ConstructionComBut As Office.CommandBarButton   '�����
Attribute ConstructionComBut.VB_VarHelpID = -1
Private WithEvents ConstructionComBut2 As Office.CommandBarButton  '��������� ����
Attribute ConstructionComBut2.VB_VarHelpID = -1

Private wallDrawer As c_WallDrawer

Private maskClass As c_Mask


Private Sub Class_Initialize()
    Set ConstructionComBut = Application.CommandBars("�����������").Controls("�����")
    Set ConstructionComBut2 = Application.CommandBars("�����������").Controls("��������� ����")
    
    Set maskClass = New c_Mask
End Sub

Private Sub Class_Terminate()
    Set ConstructionComBut = Nothing
    Set ConstructionComBut2 = Nothing
    
    Set maskClass = Nothing
End Sub

Private Sub ConstructionComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "��������� �����"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    maskClass.s_makeMask
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
    MsgBox Err.description, , ThisDocument.Name
End Sub

Private Sub ConstructionComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "����������� ����"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� ��������� ������
    If Ctrl.State = msoButtonDown Then
        wallDrawer.Term
        Set wallDrawer = Nothing
        Set visApp = Nothing
    Else
    '---���������� �����������
        Set wallDrawer = New c_WallDrawer
    '---�������� ������, ���� ����������� �������������
        If wallDrawer.Init Then Ctrl.State = msoButtonDown
    '---���������� ����� �������� �� �����������
        Set visApp = Application
    End If

Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub visApp_BeforeShapeDelete(ByVal Shape As IVShape)
'---��������� �� ���� �� ������� ������ ����� ��������� � �������� ������� ��� ������������
    If Shape = wallDrawer.wallShape Then
        wallDrawer.Term
        Set wallDrawer = Nothing
        Set visApp = Nothing
    End If
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("�����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\WallMask1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallMask2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("��������� ����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\WallDrawTool1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\WallDrawTool2.bmp")
    End With
    
Set cb = Nothing
End Sub
