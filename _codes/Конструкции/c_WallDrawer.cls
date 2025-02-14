VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_WallDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
'----------------------����� ��� ����������� ����-----------------------------------------
'-----------------------------------------------------------------------------------------
Public WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1
Public WithEvents wallShape As Visio.Shape
Attribute wallShape.VB_VarHelpID = -1
Public newWallShape As Visio.Shape
Public WallDrawerToolButton As Office.CommandBarButton  '��������� ����



'--------------------������� ����� � ������� ������--------------------------------
Public Function Init() As Boolean
    Set visApp = Application
    Set wallShape = GetSelectedWall
    Set WallDrawerToolButton = Application.CommandBars("�����������").Controls("��������� ����")
    
    If wallShape Is Nothing Then
        Init = False
        Term
    Else
        Init = True
    End If
End Function
Public Function Term() As Boolean
    Set visApp = Nothing
    WallDrawerToolButton.State = msoButtonUp
    Term = True

End Function







'Private Sub visApp_MouseDown(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
'
'    '---���������, ������� �� ���������� �����
''    If Not LineToolIsSelected Then Exit Sub
'    '---��������� ��� ��������� ������ ����� � ��� ������
''    If Not ShapeIsLine(Shape) Then Exit Sub
'
'    '---���������� ������ �����
'    Application.EventsEnabled = False
'        Set newWallShape = Application.ActivePage.Drop(wallShape, x, y)
'    Application.EventsEnabled = True
'
'    '---��������� �������� � ��������� ����� ����� �����
'    With newWallShape
'        .Cells("BeginX").Formula = x
'        .Cells("BeginY").Formula = y
''        .Cells("EndX").Formula = Shape.Cells("EndX").Formula
''        .Cells("EndY").Formula = Shape.Cells("EndY").Formula
'    End With
'End Sub
'
'
'
'Private Sub visApp_MouseMove(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
'    '---���������, ��������� �� ������ ����� �����
'    If newWallShape Is Nothing Then Exit Sub
'
'    '---��������� �������� � ��������� ����� ����� �����
'    With newWallShape
'        .Cells("EndX").Formula = x
'        .Cells("EndY").Formula = y
'    End With
'
'    DoEvents
'End Sub
'
'Private Sub visApp_MouseUp(ByVal Button As Long, ByVal KeyButtonState As Long, ByVal x As Double, ByVal y As Double, CancelDefault As Boolean)
'    Set newWallShape = Nothing
'End Sub


Private Sub visApp_ShapeAdded(ByVal Shape As IVShape)
'����� ���������� ������ ����� ������ ������ �����
Dim shp As Visio.Shape

    On Error GoTo EX

    '---���������, ������� �� ���������� �����
    If Not LineToolIsSelected Then Exit Sub
    '---��������� ��� ��������� ������ ����� � ��� ������
    If Not ShapeIsLine(Shape) Then Exit Sub

    '---���������� ������ �����
    Application.EventsEnabled = False
        Set shp = Application.ActivePage.Drop(wallShape, 0, 0)
    Application.EventsEnabled = True

    '---��������� �������� � ��������� ����� ����� �����
    With shp
        .Cells("BeginX").Formula = Shape.Cells("BeginX").Formula
        .Cells("BeginY").Formula = Shape.Cells("BeginY").Formula
        .Cells("EndX").Formula = Shape.Cells("EndX").Formula
        .Cells("EndY").Formula = Shape.Cells("EndY").Formula
    End With

    '---Delete Line and clear selection
    Shape.Delete
    '---������� ���������
    Application.ActiveWindow.DeselectAll
Exit Sub
EX:
    Application.EventsEnabled = True

    On Error Resume Next
    '---������� �����
    'Shape.Delete

    '---������� ���������
    Application.ActiveWindow.DeselectAll
End Sub

Private Function LineToolIsSelected() As Boolean
'���������� ������, ���� ���������� "�����" ������
Dim cbr As Office.CommandBar
Dim btnTool As Office.CommandBarControl
    
    On Error GoTo EX
    
    Set cbr = Application.CommandBars("Drawing")
    Set btnTool = GetCommandBarTool(cbr, 6509)
    LineToolIsSelected = btnTool.State
Exit Function
EX:
    LineToolIsSelected = False
End Function



'----------------------������ �� ������----------------------------------
Private Function GetSelectedWall() As Visio.Shape
'������� ���������� ������ ��������� �����, ���� ������� �� ����� - ���������� nothing
Dim shp As Visio.Shape

    Set shp = Application.ActiveWindow.Selection(1)
    
    If PFB_isWall(shp) Then
        Set GetSelectedWall = shp
    Else
        Set GetSelectedWall = Nothing
    End If
End Function

Private Sub wallShape_BeforeShapeDelete(ByVal Shape As IVShape)
    Term
End Sub


