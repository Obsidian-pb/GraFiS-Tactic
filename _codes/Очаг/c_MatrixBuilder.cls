VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MatrixBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------����� ��� ��������� �������----------------------------------------------------------------------------------
Option Explicit


Private grain As Integer        '������ ����� � ��
Const mmInInch As Double = 25.4
Private grain_i As Double
Private currentCells As Collection

Private form As Object

Private bordreRect As Visio.Shape

Public maskBuilder As c_MaskBuilder

Public matrix As Variant




Public Sub SetForm(ByRef frm As Object)
'��������� ������ �� ������� ����� � ������� ����� ���������� ����� � �������� ��������� �����
    Set form = frm
End Sub


Public Function NewMatrix(ByVal grainVal As Integer) As Variant()
'������� ������� � ���������� ����� ������� ������� ��� �������
'grain - ������ ����� � ��
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer
Dim areaBuilder As C_MatrixAreaBuilder
Dim shp As Visio.Shape
Dim perc As Double                      '������� �������� �����
Dim PinX As Long
Dim PinY As Long

    grain = grainVal
    grain_i = grain / mmInInch

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
    Debug.Print "������ ������� " & xCount * yCount
    
    ReDim matrix(xCount, yCount)
    
    '��������� ������� ��������� ������
    TryGetShape bordreRect, "User.IndexPers:1001"
    
    
    '�������� ���� ������� � ����� ������ ���� �������
    form.lblMatrixIsBaked.ForeColor = vbBlack
    
    '---��������� ������� ������
    FillMatrixZero

    '---��������� ������� ��������� ������������� � ������, ���� ������� ��������� ����
    If Not bordreRect Is Nothing Then
        Set areaBuilder = New C_MatrixAreaBuilder
            form.lblMatrixIsBaked.Caption = "������ ��������� ����, ���������"
'            form.lblMatrixIsBaked.ForeColor = vbBlack
            DoEvents

            FillMatrixWalls
            matrix = areaBuilder.SetForm(form).FillArea(matrix, bordreRect, grain)
        Set areaBuilder = Nothing
    End If

    '---����������� ������������� ������� ����
    form.lblMatrixIsBaked.Caption = "����������� ������� ����"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
'                Application.ActiveWindow.Select shp, visSelect
                Set currentCells = New Collection
                PinX = CellVal(shp, "PinX", visMillimeters)
                PinY = CellVal(shp, "PinY", visMillimeters)
'                Patch PinX, PinY
                Attack shp, PinX, PinY
                
                form.AddCheckedSize 1, 2
            End If
        Next shp
    Else
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
                If bordreRect.SpatialRelation(shp, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                    Set currentCells = New Collection
'                    Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters)
                    PinX = CellVal(shp, "PinX", visMillimeters)
                    PinY = CellVal(shp, "PinY", visMillimeters)
'                    Patch PinX, PinY
                    Attack shp, PinX, PinY
                    
                    form.AddCheckedSize 1, 2
                End If
            End If
        Next shp
    End If

    
    '---����������� ������������� ������� ������
    form.lblMatrixIsBaked.Caption = "����������� ������� ����"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isDoor(shp) Then
'                Application.ActiveWindow.Select shp, visSelect
                perc = CloseDoor(shp)
                
                Set currentCells = New Collection
                Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters), csOpenSpace
                
                OpenDoor shp, perc
                
                form.AddCheckedSize 1, 3
            End If
        Next shp
    Else
        For Each shp In Application.ActivePage.Shapes
            If PFB_isDoor(shp) Then
                If bordreRect.SpatialRelation(shp, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                    perc = CloseDoor(shp)
                    
                    Set currentCells = New Collection
                    Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters), csOpenSpace
                    
                    OpenDoor shp, perc
                    
                    form.AddCheckedSize 1, 3
                End If
            End If
        Next shp
    End If


    '---���������� ��������� - ������� ����������� ������� � ����������� ������������
    NewMatrix = matrix

End Function

Public Sub Attack(ByRef shp As Visio.Shape, ByVal x As Double, ByVal y As Double, Optional ByVal state As CellState = CellState.csWall)
    
Dim radiuss As Double
Dim x_i As Integer
Dim y_i As Integer
Dim pnt As c_Point

    On Error GoTo ErrRange
    
    ' ���������� � ����������� �������
    x = Int(x / grain) * grain + grain / 2
    y = Int(y / grain) * grain + grain / 2
    
    x_i = x / mmInInch
    y_i = y / mmInInch
'    grain_i = grain / mmInInch
    
    If x < 0 Or y < 0 Or x > UBound(matrix, 1) * grain Or y > UBound(matrix, 2) * grain Then
        Exit Sub
    End If
    
'    DoEvents
    
    radiuss = grain_i * 0.75
'    radiuss = grain_i

    Set pnt = New c_Point
    pnt.SetData Int(x / grain), Int(y / grain)
    If IsInCollection(currentCells, pnt) Then
        Exit Sub
    End If

'                '������������ �������
'                Dim shpT As Visio.Shape
'                Set shpT = Application.ActivePage.DrawRectangle(x / mmInInch, y / mmInInch, _
'                        (x + grain) / mmInInch, (y + grain) / mmInInch)
'                shpT.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("test")
'                shpT.Cells("LinePattern").FormulaU = "0"
'                shpT.Cells("FillForegnd").FormulaU = "2"
    If shp.HitTest(x_i, y_i, radiuss) Then
'        Debug.Print "c_MatrixBuilder->Attack->shp.HitTest: �������� ����� � ������� ������ - �� ����������� ��� ������ ����"
        
        currentCells.Add pnt
        matrix(Int(x / grain), Int(y / grain)) = state
        
        Attack shp, x - grain, y, state
        Attack shp, x - grain, y + grain, state
        Attack shp, grain, y + grain, state
        Attack shp, x + grain, y + grain, state
        Attack shp, x + grain, y, state
        Attack shp, x + grain, y - grain, state
        Attack shp, grain, y - grain, state
        Attack shp, x - grain, y - grain, state
    End If
    
Exit Sub
ErrRange:
    If Err.number = 9 Then
        Exit Sub
    Else
        Debug.Print Err.description
    End If
End Sub

Private Sub CheckEnvironment(Optional ByRef rect As c_Rect)
'��������� ���������� ������������ ������������
Dim startRectSizeHalf As Long
    
    '���������� ����������� �����
    '---���������� ������ ���������� �������� ��� ������ ����
        startRectSizeHalf = GetMaximumRectSize / 2   '�������� ������ ��� � ����� ������ � ������ � ���������� ����� �����
        '������ ��� ������� ����� �������� ������� �������
        form.SetMatrixSize (startRectSizeHalf * 2) ^ 2
    '---��������� ����������� ������
        CheckRect startRectSizeHalf, startRectSizeHalf, startRectSizeHalf
End Sub


Private Function IsOpenSpace(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ����, � ��������, ���� �������������
    IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function

Private Function IsTotalWall(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ���������� ������������, � ��������, ���� �������������
    IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function

Private Sub Patch(x, y, Optional ByVal state As CellState = CellState.csOpenSpace)
' ��������� ��������� � ����� ������� � ������������ x, y �������� �� ������ � ������� ����������� �������� CellState
Dim xx As Integer
Dim yy As Integer

    xx = Int(x / grain)
    yy = Int(y / grain)

'                '������������ �������
'                Dim shpT As Visio.Shape
'                Set shpT = Application.ActivePage.DrawRectangle(x / mmInInch, y / mmInInch, _
'                        (x + grain) / mmInInch, (y + grain) / mmInInch)
'                shpT.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("test")
'                shpT.Cells("LinePattern").FormulaU = "0"
'                shpT.Cells("FillForegnd").FormulaU = "2"

    matrix(xx - 1, yy - 1) = state
    matrix(xx - 1, yy) = state
    matrix(xx - 1, yy + 1) = state
    matrix(xx, yy - 1) = state
    matrix(xx, yy) = state
    matrix(xx, yy + 1) = state
    matrix(xx + 1, yy - 1) = state
    matrix(xx + 1, yy) = state
    matrix(xx + 1, yy + 1) = state
End Sub




'------------------����������� ����������� �������------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'����� radiuss - � ������!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       '����������� ������� - �����, ��� �� ������������� �������� ���� ������� �������� - ��������� ����� ���������� ������� (����������� ��������� ����� ��������) � ������� ����� ������������ �������, � ������ ����� 1,4142135623731

    On Error GoTo EX
               
               
'1 ��������� �����������, ��� � ������� ��� ����
    '---���� ���� ���, ��������� ��� ���� ������ 0
        If IsOpenSpace(x, y, radiuss * radiussModificator * grain) Then
            form.AddCheckedSize (radiuss * 2) ^ 2

'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                    (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'            shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
            DoEvents

            Exit Sub
        End If
        
'2 ���� ����� ����, ��������� ������� ��������
'3 ��������� �����������, ��� � ������� ��� ��������� ������������
    '---���� ����� ����, ��������� ����� ��������
        If IsTotalWall(x, y, radiuss * radiussModificator * grain) Then
        '���� ��� ���� � ������� ��� ��������� ������������, ������ ��� ���� �������� ����� �������� ����� (1)
            SetToAllCells x, y, radiuss, csWall
            form.AddCheckedSize (radiuss * 2) ^ 2
            
'                Dim shp As Visio.Shape
'                Set shp = Application.ActivePage.DrawRectangle((X - radiuss) * grain / mmInInch, (Y - radiuss) * grain / mmInInch, _
'                        (X + radiuss) * grain / mmInInch, (Y + radiuss) * grain / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
                DoEvents
        Else
            If radiuss = 1 Then
                form.AddCheckedSize 1
                
                '��������� ��� �������� ������, �.�. ��� ������� = 1 � ������ ������ ������ ���������
'                radiussHalf = 0.5
'                '1 - ����� �����
'                    checkSingleCell x - 0.5, y + 0.5
'                '2 - ����� ������
'                    checkSingleCell x + 0.5, y + 0.5
'                '3 - ���� ������
'                    checkSingleCell x + 0.5, y - 0.5
'                '4 - ���� �����
'                    checkSingleCell x - 0.5, y - 0.5
                '1 - ����� �����
                    CheckSingleCell x, y, False, True
                '2 - ����� ������
                    CheckSingleCell x, y, True, True
                '3 - ���� ������
                    CheckSingleCell x, y, True, False
                '4 - ���� �����
                    CheckSingleCell x, y, False, False
'                ��������� ������� ������, �.�. ��� ������� = 0 ��� �������� ��������� � ����� �� �������
'                checkSingleCell x + 0.5, y + 0.5
                Exit Sub
'            End If
            Else
            
            '����� ���������� �������� �������� ������� �������� ���������
            radiussHalf = radiuss * 0.5
            '1 - ����� �����
                CheckRect x - radiussHalf, y + radiussHalf, radiussHalf
            '2 - ����� ������
                CheckRect x + radiussHalf, y + radiussHalf, radiussHalf
            '3 - ���� ������
                CheckRect x + radiussHalf, y - radiussHalf, radiussHalf
            '4 - ���� �����
                CheckRect x - radiussHalf, y - radiussHalf, radiussHalf
            
            End If
        End If
Exit Sub
EX:
    Debug.Print "x=" & x & ", y=" & y & ", radiuss=" & radiuss
End Sub



Private Sub CheckSingleCell(ByVal x As Integer, ByVal y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean, Optional tolerance As Double = 0)
'��������� ������� ���� ��� ��������� (��������� ������)
Dim xAct As Double
Dim yAct As Double
Const radiussModificator As Double = 1.4142

    xAct = x + 0.5 + Not dirRight
    yAct = y + 0.5 + Not dirUp
    
    If Not maskBuilder.pO_MaskShape.HitTest(xAct * grain / mmInInch, yAct * grain / mmInInch, radiussModificator * grain * 0.5 / mmInInch) = 0 Then
        If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csWall
        If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csWall
        If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csWall
        If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csWall
        
        
'        Debug.Print x * grain & ":" & y * grain & " �����!"
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    Else
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    End If

End Sub

Private Sub SetToAllCells(ByVal x As Long, ByVal y As Long, ByVal size As Long, ByVal value As CellState)
'������������� ��� ���� ������ ���������� ��������� size � ������� � ����� x,y �������� value
Dim i As Long
Dim j As Long
    
    For i = x - size To x + size
        For j = y - size To y + size
            matrix(i, j) = value
        Next j
    Next i
End Sub

Private Function GetMaximumRectSize() As Long
'���������� ����������� ��������� ������� ��������� ����������� ����� � ��� ���� ���������� ��� ������� ����� ��������� ��������� �� 2
Dim grains As Long
Dim pageWidth As Double
Dim pageHeight As Double

    grains = 1
    
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)
    
    Do While grains * grain < pageWidth Or grains * grain < pageHeight
        grains = grains * 2
    Loop
    
GetMaximumRectSize = grains
End Function

'------------------����������� ����������� �������------------------------------------------------------

Private Sub FillMatrixZero()
'��������� ������� ������
Dim x As Long
Dim y As Long
    
    For y = 0 To UBound(matrix, 1)
        For x = 0 To UBound(matrix, 2)
            matrix(y, x) = CellState.csOpenSpace
        Next x
    Next y
End Sub

Private Sub FillMatrixWalls()
'��������� ������� ������
Dim x As Long
Dim y As Long
    
    For y = 0 To UBound(matrix, 1)
        For x = 0 To UBound(matrix, 2)
            matrix(y, x) = CellState.csWall
        Next x
    Next y
End Sub

'------------------�������-----------------------
Private Function CloseDoor(ByRef shp As Visio.Shape) As Double
Dim perc As Double
    
    CloseDoor = CellVal(shp, "Prop.VisDoorOpenPercent")
    SetCellVal shp, "Prop.VisDoorOpenPercent", 0
    
End Function

Private Sub OpenDoor(ByRef shp As Visio.Shape, Optional ByVal perc As Double = 50)
     SetCellVal shp, "Prop.VisDoorOpenPercent", perc
End Sub
