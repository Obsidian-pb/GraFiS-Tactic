VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_WorkPlaces"
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
Private col_PlaceShapes As Collection '��������� ���������� ����� ������� ����
Private col_WallsShapes As Collection '��������� ����




'--------------------------��������� ������-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set vO_App = Visio.Application
End Sub

Private Sub Class_Terminate()
    Set vO_App = Nothing
End Sub

Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
    col_RoomsShapes.Add Shape
End Sub


'-----------------------------------�������� ���������------------------------------------------------------------------------

Public Sub S_SetFullShape()
'�������� ��������� ������� ������ ������� ������� ���� �� �����
Dim vO_SheetShape As Visio.Shape


On Error GoTo Tail

    '---���������� ��������� �����
        Set col_RoomsShapes = New Collection
        Set col_PlaceShapes = New Collection
        Set col_WallsShapes = New Collection
    
    '---��������� ������������ ������� ���������� �� ������� (����� ����� ������)
        vO_App.EventsEnabled = False
    
    '---��������� �����
        s_makeMask
        
    '---��������� ������� �� ������������ �����, ���� ��� - �������
        If pO_MaskShape Is Nothing Then
            vO_App.EventsEnabled = True
            Set col_RoomsShapes = Nothing
            Set col_PlaceShapes = Nothing
            Set col_WallsShapes = Nothing
            Exit Sub
        End If
    
    '---���������� ������ ������ � ��������� �����
        Set vO_SheetShape = Application.ActiveWindow.Page.DrawRectangle(0, 0, _
                        Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches), _
                        Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches))
    
    '---�������� �� ������ ����� �������
        Application.ActiveWindow.DeselectAll
            Application.ActiveWindow.Select vO_SheetShape, visSelect
            Application.ActiveWindow.Select pO_MaskShape, visSelect
        Application.ActiveWindow.Selection.Subtract
        '---�������� ������������ ������� ���������� �� ������� (����� ����� ������)
        vO_App.EventsEnabled = True
        '---������������� ���������� ������
        Application.ActiveWindow.Selection.Fragment
    
    '---����������� ������� ����� (������� ������ � ������ ���� - � ���� ��� ������ ����� ����������� ������)
        ps_DeleteTrash
       
    '---�������� ���������� ������� �������� ����
        '---��������� ������������ ������� ���������� �� ������� (����� ����� ������)
        vO_App.EventsEnabled = False
        '��������� ������ ������� ����
        ps_InsertWorkPlaces
        '������� ��� ������ ������
        ps_DeleteRooms
    
    '---�������� ������������ ������� ���������� �� ������� (����� ����� ������)
    vO_App.EventsEnabled = True
    Set col_RoomsShapes = Nothing
    Set col_PlaceShapes = Nothing
    Set col_WallsShapes = Nothing
Exit Sub
Tail:
'    Debug.Print Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "S_SetFullShape"
    '---�������� ������������ ������� ���������� �� ������� (����� ����� ������)
    vO_App.EventsEnabled = True
    Set col_RoomsShapes = Nothing
    Set col_PlaceShapes = Nothing
    Set col_WallsShapes = Nothing
End Sub

Private Sub ps_DeleteTrash()
'����� ������� ������ ��������
Dim vO_Shape As Visio.Shape
Dim tmpCol As Collection

    Set tmpCol = New Collection
    
    '������� ������� � ������������ ��������� ��������� ��������� ������
    For Each vO_Shape In col_RoomsShapes
        If PFI_FirstSectionCount(vO_Shape) > 1 Then
            vO_Shape.Delete
        Else
            tmpCol.Add vO_Shape
        End If
    Next vO_Shape

    '---��������� ��������� ������
    Set col_RoomsShapes = tmpCol
End Sub

Private Sub ps_InsertWorkPlaces()
'����� �������� ��� ������ ������ �������� ������� ����
Dim vO_RoomShape As Visio.Shape
Dim vO_PlaceShape As Visio.Shape
Dim vO_Master As Visio.Master
Dim i As Integer
    
'    Set vO_Master = Application.Documents("WALL_M.vss").Masters(4) '("�����")
    Set vO_Master = Application.Documents("WALL_M.vss").Masters("�����")
    
On Error Resume Next
    
    i = 1
    '---���������� ��� ������ � ���������������
    For Each vO_RoomShape In col_RoomsShapes
        '���������� ������ �������� �����
        Set vO_PlaceShape = Application.ActivePage.Drop(vO_Master, 0, 0)
        '������ �� ��������� �������
        PF_GeometryCopy vO_RoomShape, vO_PlaceShape
        '��������� ������� ������
        vO_PlaceShape.Cells("User.visFMArea").FormulaU = str(vO_PlaceShape.AreaIU)
        '���������� ������ �� ������ ����
'        vO_PlaceShape.SendToBack
        '��������� ������ ����� � ��������� � ����������� �� ��������� ���������� �����
        col_PlaceShapes.Add vO_PlaceShape
        vO_PlaceShape.Cells("Prop.LocationID").FormulaU = i
        
        i = i + 1
    Next vO_RoomShape

End Sub

Private Sub ps_DeleteRooms()
'����� ������� ��� ������ ������ � ������� ���������
Dim vO_RoomShape As Visio.Shape

For Each vO_RoomShape In col_RoomsShapes
    vO_RoomShape.Delete
    col_RoomsShapes.Remove 1
Next vO_RoomShape

End Sub

Public Sub PF_GeometryCopy(ByRef a_RoomShape As Visio.Shape, ByRef a_PlaceShape As Visio.Shape)
'����� ������� ������ �������� ����� ��������� ������ �������
Dim j As Integer
Dim k As Integer
    
    On Error GoTo Tail
    
    '---��������� ���������
    '---������� ��������� �����
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
        a_PlaceShape.DeleteRow visSectionFirstComponent, 1
    '---�������� ��������� �� �������
        j = 1
        Do While a_RoomShape.RowExists(visSectionFirstComponent, j, 0)
            a_PlaceShape.AddRow visSectionFirstComponent, j, a_RoomShape.RowType(visSectionFirstComponent, j)
            
            k = 0
            Do While a_RoomShape.CellsSRCExists(visSectionFirstComponent, j, k, 0)
                a_PlaceShape.CellsSRC(visSectionFirstComponent, j, k).FormulaU = _
                    a_RoomShape.CellsSRC(visSectionFirstComponent, j, k).FormulaU
                
                k = k + 1
            Loop
            j = j + 1
        Loop
    '---������������ ��������� � ������� �������� ������ � ������ �������
        a_PlaceShape.Cells("Width").FormulaU = a_RoomShape.Cells("Width").FormulaU
        a_PlaceShape.Cells("Height").FormulaU = a_RoomShape.Cells("Height").FormulaU
        a_PlaceShape.Cells("LocPinX").FormulaU = a_RoomShape.Cells("LocPinX").FormulaU
        a_PlaceShape.Cells("LocPinY").FormulaU = a_RoomShape.Cells("LocPinY").FormulaU
        a_PlaceShape.Cells("PinX").FormulaU = a_RoomShape.Cells("PinX").FormulaU
        a_PlaceShape.Cells("PinY").FormulaU = a_RoomShape.Cells("PinY").FormulaU

Exit Sub
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PF_GeometryCopy"
End Sub



'-----------------------------------------------------��������� �������� �����----------------------------------------------------

Private Sub s_makeMask()
'��������� ������� ����� ������� (��������)
Dim vO_Shape As Visio.Shape

    '---��������� ���������
    For Each vO_Shape In Application.ActivePage.Shapes
        ps_AddWall vO_Shape
    Next vO_Shape

    '---������ ����� ����
   s_makeMaskWalls

Set vO_Shape = Nothing
End Sub



Private Sub s_makeMaskWalls()
'��������� ������� ����� ������� ����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim X1 As Double, Y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---��������� ������� �� ������ � ��������� col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub

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
    Application.ActiveWindow.Selection.SendToBack

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


'-----------------------------------------��������� ����������� ���������-------------------------------------------------------
Private Sub ps_AddWall(ByRef aO_Shape As Visio.Shape)
'��������� ���������� ���������� ��� ������ ����������� �� ����� � ��������� � ��������� col_WallsShapes ����� �����������
'� �������� ��������� �������
Dim vO_shp As Visio.Shape

'---���������, �������� �� ������ ��������� (���������������)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            ps_AddWall vO_shp
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---���������, �������� �� ������ ������� �����
    If PFB_isWall(aO_Shape) Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If

End Sub


'---------------------------------------��������� ������� � �����--------------------------------------------------
Private Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'������� �������������� ������������ ������������� ��������
    If Shape.Parent.Name = Application.ActivePage.Name Then
        AngleToPage = Shape.Cells("Angle")
    Else
        AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
    End If

Set Shape = Nothing
End Function


