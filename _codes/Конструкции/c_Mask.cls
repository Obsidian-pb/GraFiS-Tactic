VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Mask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Private pO_MaskShape As Visio.Shape '������ ����� - ����������� �� ����� ���� � ������������ ��� ����� ���� ���� ��� �������

'--------------------------���������--------------------------------------------------------------------------------------------
Private col_RoomsShapes As Collection '��������� ���������� ����� ������
'Private col_PlaceShapes As Collection '��������� ���������� ����� ������� ����
Private col_WallsShapes As Collection '��������� ����
Private col_DoorsShapes As Collection '��������� ������
Private col_WindowsShapes As Collection '��������� ����



'--------------------------��������� ������-------------------------------------------------------------------------------------
'Private Sub Class_Initialize()
'    Set vO_App = Visio.Application
'End Sub
'
'Private Sub Class_Terminate()
'    Set vO_App = Nothing
'End Sub

'Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
'    col_RoomsShapes.Add Shape
'End Sub



'-----------------------------------------------------��������� �������� �����----------------------------------------------------

Public Sub s_makeMask()
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
    
    '---������� ������ � ���� �����
    ClearLayer "�����"
    
    '---��������� ���������
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
    Set col_WindowsShapes = New Collection
    
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddShape vO_Shape
    Next vO_Shape
    
    '---������ ����� ����
    '---��������� ������� �� ������ � ��������� col_WallsShapes
    If col_WallsShapes.count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---������� �� ����� ���� ������
    If col_WallsShapes.count = 0 And col_DoorsShapes.count = 0 And col_WindowsShapes.count = 0 Then Exit Sub
    s_deleteWindowsAndDoors


Set vO_Shape = Nothing
End Sub


Private Sub s_makeMaskWalls()
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim X1 As Double, Y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---���������� ��� ������ ����
    For Each vO_Shape In col_WallsShapes
        '---���������� ���������� ��� ������
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), X1, Y1
        '---���������� ����� ������ �� �����������
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), X1, Y1)
            vO_TempShape.Cells("PinX") = X1
            vO_TempShape.Cells("PinY") = Y1
            vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("LinePattern") = 0
            vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
            
        '---��������� �� ��������� ���������
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---��������� �� ������ ��������� ��������� �����
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Union

'---���������� �����
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

    Set col_Shadows = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "s_makeMaskWalls"
    Set col_Shadows = Nothing
End Sub

Private Sub s_deleteWindowsAndDoors()
'����� ������� �� ����� ���� �� ������ - ���� � �����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_WindowsAndDoors As Collection
Dim X1 As Double, Y1 As Double
Dim rect As c_Rect
    
    Set col_WindowsAndDoors = New Collection
    
    '��������� ��������� ��������������� ��� ���������� �� ����� ����
    '---���������� ��� ������ ����
    For Each vO_Shape In col_WindowsShapes
        Set rect = New c_Rect                       '������� ����� ������ ��������������
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_WindowsAndDoors.Add vO_TempShape
    Next vO_Shape
    '---���������� ��� ������ ������
    For Each vO_Shape In col_DoorsShapes
        Set rect = New c_Rect                       '������� ����� ������ ��������������
        Set vO_TempShape = rect.GetRectShape(vO_Shape)
        col_WindowsAndDoors.Add vO_TempShape
    Next vO_Shape
    
    '---������� ���������
    Application.ActiveWindow.DeselectAll
    '---��������� � �������������� ��������� ����� ����
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---���������� ��� ������ ��������������� � ��������� �� � ���������
    For Each vO_Shape In col_WindowsAndDoors
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    
    '---������� ������ �� �����
    Application.ActiveWindow.Selection.Subtract
    
    '---���������� �����
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)
    
    '---��������� ����� � ���� "����� ����"
    pO_MaskShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("�����") & """"
End Sub





'-----------------------------------------��������� ����������� ���������-------------------------------------------------------
Private Sub ps_AddShape(ByRef aO_Shape As Visio.Shape)
'��������� ���������� ���������� ��� ������ ����������� �� ����� � ��������� � ��������� col_WallsShapes ����� �����������
'� �������� ��������� �������
Dim vO_shp As Visio.Shape

'---���������, �������� �� ������ ��������� (���������������)
    If aO_Shape.Shapes.count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddShape vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---���������, �������� �� ������ ������� �����
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If
'---���������, �������� �� ������ ������� ����
    If PFB_isWindow(aO_Shape) Then
        col_WindowsShapes.Add aO_Shape
        Exit Sub
    End If
'---���������, �������� �� ������ ������� �����
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub






