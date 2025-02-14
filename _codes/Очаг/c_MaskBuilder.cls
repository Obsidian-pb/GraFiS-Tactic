VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MaskBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------����� ������� � ���������� ������ ����� ����-------------------------------
Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Public pO_MaskShape As Visio.Shape                 '������ ����� - ����������� �� ����� ���� � ������������ ��� ����� ���� ���� ��� �������
Public pO_FreespaceMaskShape As Visio.Shape        '������ ����� ���������� ������������ - ���������� ����������� �� �������������� �������� ����� ����
Const mmInInch As Double = 25.4


'--------------------------���������--------------------------------------------------------------------------------------------
Private col_WallsShapes As Collection '��������� ����
Private col_DoorsShapes As Collection '��������� ������
Private col_WindowsShapes As Collection '��������� ����







'-----------------------------------------------------��������� �������� �����----------------------------------------------------
Public Sub s_makeMask(Optional ByRef borders As Visio.Shape = Nothing)
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
    
    '---������� ������ � ���� �����
    ClearLayer "�����"
    
    '---��������� ���������
    Set col_WallsShapes = New Collection
    Set col_DoorsShapes = New Collection
    
    
    If borders Is Nothing Then
        For Each vO_Shape In Application.ActivePage.Shapes
            ps_AddShape vO_Shape
        Next vO_Shape
    Else
        ' �������� �� SpatialSearch
        For Each vO_Shape In Application.ActivePage.Shapes
            If borders.SpatialRelation(vO_Shape, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                ps_AddShape vO_Shape
            End If
        Next vO_Shape
    End If
    
    '---������ ����� ����
    '---��������� ������� �� ������ � ��������� col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub
    s_makeMaskWalls
    
    '---������� �� ����� ���� ������
    If col_WallsShapes.Count = 0 And col_DoorsShapes.Count = 0 Then Exit Sub
    s_deleteDoors
    
    '---��������� � ����� ���� ������
    

Set vO_Shape = Nothing
End Sub

Public Sub s_makeOpenSpaceMask()
'��������� ������� ����� ��������� ������������
Dim pageShape As Visio.Shape
Dim pO_MaskShapeClone As Visio.Shape
Dim pageWidth As Double
Dim pageHeight As Double
    
    '---���������� ������ ��������
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)

    '---������ ������������� �������� �� ��������
    Set pageShape = Application.ActivePage.DrawRectangle(0, 0, pageWidth / mmInInch, pageHeight / mmInInch)
    
    '---��������� ���� ����� ����
    Dim x1 As Double, y1 As Double
    '---���������� ���������� ��� ������
    pO_MaskShape.XYToPage pO_MaskShape.Cells("LocPinX").Result(visInches), pO_MaskShape.Cells("LocPinY").Result(visInches), x1, y1
    '---���������� ����� ������ �� �����������
    Set pO_MaskShapeClone = Application.ActivePage.Drop(pO_MaskShape, x1, y1)
       
    '---��������� � ��������� ������������� �������� � ����� ����
    Application.ActiveWindow.Selection.DeselectAll
    Application.ActiveWindow.Select pageShape, visSelect
    Application.ActiveWindow.Select pO_MaskShapeClone, visSelect
    '---�������� �� �������������� ����� ���� ����� ����
    Application.ActiveWindow.Selection.Subtract

    '---���������� �����
    Set pO_FreespaceMaskShape = Application.ActiveWindow.Selection(1)
    
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillForegnd).FormulaU = "THEMEGUARD(TINT(THEME(""FillColor""),24))"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillBkgnd).FormulaU = "THEMEGUARD(SHADE(FillForegnd,LUMDIFF(THEME(""FillColor""),THEME(""FillColor2""))))"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowLine, visLineColorTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillForegndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillShdwForegndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowFill, visFillShdwBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowText, visTxtBlkBkgndTrans).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionObject, visRowImage, visImageTransparency).FormulaU = "90%"
    pO_FreespaceMaskShape.CellsSRC(visSectionCharacter, 0, visCharacterColorTrans).FormulaU = "90%"

End Sub

Private Sub s_makeMaskWalls()
'��������� ������� ����� ����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double
Dim rect As c_Rect

Set col_Shadows = New Collection

On Error GoTo Tail

'---���������� ��� ������ ����
    For Each vO_Shape In col_WallsShapes
        '---���������� ���������� ��� ������
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
        '---���������� ����� ������ �� ����������� (�������� ��� ���� � ������)
            If vO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then  '�����
                Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
                
                vO_TempShape.Cells("PinX") = x1
                vO_TempShape.Cells("PinY") = y1
                vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
                vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
                vO_TempShape.Cells("LinePattern") = 0
                vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
                '---��������� �� ��������� ���������
                col_Shadows.Add vO_TempShape
            ElseIf vO_Shape.Cells("User.ShapeType").Result(visNumber) = 6 Then  '�������
                If F_InsertFire.flag_ColumnsCalculate.value = True Then
                    Set vO_TempShape = GetColumnShadow(vO_Shape)

                    '---��������� �� ��������� ���������
                    col_Shadows.Add vO_TempShape
                End If
            End If
            
            
            
        
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
    Debug.Print Err.description & " -> " & Err.number
    Set col_Shadows = Nothing
End Sub

Private Sub s_deleteDoors()
'����� ������� �� ����� ����  �����
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Doors As Collection
Dim x1 As Double, y1 As Double
Dim rect As c_Rect
    
    Set col_Doors = New Collection
    
    On Error GoTo Tail
    
    '��������� ��������� ��������������� ��� ���������� �� ����� ����
    '---���������� ��� ������ ������
    For Each vO_Shape In col_DoorsShapes
        '���������� ���������� �� ��������� �������� ����� (�� ����� �� ���� �������)
            '��������� ������� �� �����
            If DoorIsOpen(vO_Shape) Then
                Set rect = New c_Rect                       '������� ����� ������ ��������������
                Set vO_TempShape = rect.GetRectShape(vO_Shape)
                col_Doors.Add vO_TempShape
            End If

    Next vO_Shape
    
    '---������� ���������
    Application.ActiveWindow.DeselectAll
    '---��������� � �������������� ��������� ����� ����
    Application.ActiveWindow.Select pO_MaskShape, visSelect
    
    '---���������� ��� ������ ��������������� � ��������� �� � ���������
    For Each vO_Shape In col_Doors
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    
    '---������� ������ �� �����
    Application.ActiveWindow.Selection.Subtract
    
    '---���������� �����
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)
    
    '---��������� ����� � ���� "����� ����"
    pO_MaskShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """" & GetLayerNumber("�����") & """"
    
    Set col_Doors = Nothing
Exit Sub
Tail:
    Set col_Doors = Nothing
End Sub

Private Function DoorIsOpen(ByRef aO_DoorShape As Visio.Shape) As Boolean
'������� ���������� True ���� ����� �������
'!!!�� ���� ����������� ������ ������ ������!!!
Dim vI_DoorType As Integer '��� �����

On Error GoTo Tail
    
    '---��������� ����� �� �� ���� � ������� ������, ���� �� - ������� �� ������� � �������� ������
    If aO_DoorShape.Cells("User.ShapeType").Result(visNumber) = 25 Then
        DoorIsOpen = True
        Exit Function
    End If
    
    '---��������� ����� �� �� ���� � ������� �����, ���� ��� - ������� �� �������
    If Not aO_DoorShape.Cells("User.ShapeType").Result(visNumber) = 10 Then
        DoorIsOpen = False
        Exit Function
    End If

    vI_DoorType = aO_DoorShape.Cells("User.SubShapeType").Result(visNumber)
    Select Case vI_DoorType
        Case Is = 74 '��������������, ������������� ���������
            If aO_DoorShape.Shapes(9).Cells("Angle") <> 0 Then
                DoorIsOpen = True
            End If
        Case Is = 28 '�������, � ������� ���������, �������������� �������
            Select Case aO_DoorShape.Cells("User.SubShapeType.Prompt").ResultStr(visUnitsString)
            Case Is = "�������"
                If aO_DoorShape.Shapes(11).Cells("Angle") <> 0 Then
                    DoorIsOpen = True
                End If
            Case Is = "�������������� �������"
                If aO_DoorShape.Shapes(9).Cells("Angle") <> 0 Then
                    DoorIsOpen = True
                End If
            Case Is = "� ������� ���������"
                If aO_DoorShape.Shapes(11).Cells("Angle") <> 0 Then
                    DoorIsOpen = True
                End If
            End Select
        Case Is = 64 '�����������
            If aO_DoorShape.Shapes(9).Cells("Angle").ResultInt(visDegrees, 1) = 0 Or _
                aO_DoorShape.Shapes(10).Cells("Angle").ResultInt(visDegrees, 1) = 0 Or _
                aO_DoorShape.Shapes(11).Cells("Angle").ResultInt(visDegrees, 1) = 0 Or _
                aO_DoorShape.Shapes(12).Cells("Angle").ResultInt(visDegrees, 1) = 0 Then
                DoorIsOpen = False
            Else
                DoorIsOpen = True
            End If
            
        '!!!��� ������ ���� ����������� ���������
        Case Is = 72 '� ��������
            DoorIsOpen = True
        Case Is = 26 '������������� � ��������
            DoorIsOpen = True
        Case Is = 73 '�������������� ����������
            DoorIsOpen = True
        Case Is = 71 '�������������� ��������������
            DoorIsOpen = True
        Case Is = 25 '������������� ��������������
            DoorIsOpen = True
        Case Is = 65 '����������
            DoorIsOpen = True
        Case Else    '��� ���� ������ �������
            DoorIsOpen = True
    End Select
    
Exit Function

Tail:
    DoorIsOpen = True
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "DoorIsOpen"
End Function

Public Function GetColumnShadow(ByRef shp As Visio.Shape) As Visio.Shape
'���������� ���� �������: 0 - ����������, 1 - �������
'!!!�� ���� ����������� ������ ������ ������!!!
Dim newShp As Visio.Shape
Dim x, y, w, h, xloc, yloc, x1, y1, x2, y2 As Double

    '����������� ��������� ����������� ������
    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)
    xloc = shp.Cells("LocPinX").Result(visInches)
    yloc = shp.Cells("LocPinY").Result(visInches)
    w = shp.Cells("Width").Result(visInches)
    h = shp.Cells("Height").Result(visInches)
    x1 = x - xloc
    y1 = y - yloc
    x2 = x + (w - xloc)
    y2 = y + (h - yloc)
    
    '��������� ��������� ������
    If shp.RowType(visSectionFirstComponent + 1, 2) = visTagLineTo Then
        Set newShp = Application.ActiveWindow.Page.DrawRectangle(x1, y1, x2, y2)
    ElseIf shp.RowType(visSectionFirstComponent + 1, 2) = visTagArcTo Then
        Set newShp = Application.ActiveWindow.Page.DrawOval(x1, y1, x2, y2)
    End If
    
    Set GetColumnShadow = newShp
End Function

'-----------------------------------------��������� ����������� ���������-------------------------------------------------------
Private Sub ps_AddShape(ByRef aO_Shape As Visio.Shape)
'��������� ���������� ���������� ��� ������ ����������� �� ����� � ��������� � ��������� col_WallsShapes ����� �����������
'� �������� ��������� �������
Dim vO_shp As Visio.Shape

'---���������, ��������� �� ������ � �������� �������� �����
    If Not IsShapeOnSheet(aO_Shape) Then
        Exit Sub
    End If

'---���������, �������� �� ������ ��������� (���������������)
    If aO_Shape.Shapes.Count > 1 Then
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
''---���������, �������� �� ������ ������� ����
'    If PFB_isWindow(aO_Shape) Then
'        col_WindowsShapes.Add aO_Shape
'        Exit Sub
'    End If
'---���������, �������� �� ������ ������� �����
    If PFB_isDoor(aO_Shape) Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If
    

End Sub


