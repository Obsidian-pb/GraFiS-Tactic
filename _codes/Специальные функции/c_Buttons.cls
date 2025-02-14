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

Public WithEvents ComBut As Office.CommandBarButton   '������� � JPG
Attribute ComBut.VB_VarHelpID = -1
Public WithEvents ComBut2 As Office.CommandBarButton  '������
Attribute ComBut2.VB_VarHelpID = -1
Public WithEvents ComBut3 As Office.CommandBarButton  '��������� ���������
Attribute ComBut3.VB_VarHelpID = -1
Public WithEvents ComBut4 As Office.CommandBarButton  '���������� �����
Attribute ComBut4.VB_VarHelpID = -1
Public WithEvents ComBut5 As Office.CommandBarButton  '������
Attribute ComBut5.VB_VarHelpID = -1
Public WithEvents ComBut6 As Office.CommandBarButton  '������� �������� (����� �� ���� ������ ������)
Attribute ComBut6.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("�����������").Controls("������� � JPG")
    Set ComBut2 = Application.CommandBars("�����������").Controls("������")
    Set ComBut3 = Application.CommandBars("�����������").Controls("��������� ������������")
    Set ComBut4 = Application.CommandBars("�����������").Controls("���������� �����")
    Set ComBut5 = Application.CommandBars("�����������").Controls("������")
    Set ComBut6 = Application.CommandBars("�����������").Controls("������� ��������")
    
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
    Set ComBut2 = Nothing
    Set ComBut3 = Nothing
    Set ComBut4 = Nothing
    Set ComBut5 = Nothing
    Set ComBut6 = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "������� � JPG"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    JPGExportAll_P
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComBut2_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "������"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    SetAspect_P
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComBut3_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "��������� ������������"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    FixZIndex_P
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComBut4_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "���������� �����"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---��������� �������� �������
    ShapesCountShow
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComBut5_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "������"
Dim ShpObj As Visio.Shape
Dim ShpInd As Integer

'---�������� ��������� ������ - ��� �������������� ������� ������ ��� ������� ��������� ������
    On Error GoTo Tail
    
'---����������/�������� ������ ������������ "������"
    If Ctrl.state = msoButtonUp Then
        AddTimer
    Else
        DelTBTimer
    End If
    
Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Private Sub ComBut6_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'����������� ������� ������� ������ "������� ��������"
Dim ShpObj As Visio.Shape
Dim ShpInd As Variant
Dim shpIndArr As Variant
Dim indexPersColl As Collection

'---�������� ��������� ������ - ��� �������������� ������� ������
    On Error GoTo Tail
    
'---�������� ��� ����������� ������
    '1 ����������� IndexPers � ����� � �������
        If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
        
        Set indexPersColl = New Collection
        
        For Each ShpObj In Application.ActiveWindow.Selection
            If IsGFSShape(ShpObj) Then
                Set indexPersColl = AddToColl(indexPersColl, ShpObj.Cells("User.IndexPers").ResultStr(0))
            End If
        Next ShpObj
    
    '2 �������� ��� ������ � ���������� IndexPers
        shpIndArr = ColToArr(indexPersColl)
        
        For Each ShpObj In Application.ActivePage.Shapes
            If IsGFSShapeWithIP(ShpObj, shpIndArr) Then
                ActiveWindow.Select ShpObj, visSelect
            End If
        Next ShpObj

Exit Sub
Tail:
    '---������� �� ��������� ���������
End Sub

Public Sub PictureRefresh()
'��������� �������� �� �������
Dim cb As Office.CommandBarButton
Dim DocPath As String
    
'---��������� ������ �� �������
    DocPath = ThisDocument.path
    Set cb = Application.CommandBars("�����������").Controls("������� � JPG")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\ExportJPG1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\ExportJPG2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("������")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Aspect1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Aspect2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("��������� ������������")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Fix1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Fix2.bmp")
    End With
    
    Set cb = Application.CommandBars("�����������").Controls("���������� �����")
    With cb
        .Picture = LoadPicture(DocPath & "Bitmaps\Count1.bmp")
        .Mask = LoadPicture(DocPath & "Bitmaps\Count2.bmp")
    End With
    
Set cb = Nothing
End Sub
