VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public fuelValueLayer As Variant                            '����� �������� - �������. ���� ����� 0 , ������ ������ � ������ �� ����������.
'Public lowerBurnBoundLayer As Variant                       '������ ������� ��� ������� ������ �������� �������������� ������� �� �������� (� ��� �� ������ �������������)
'Public maximumBurnPowerLayer As Variant                     '������������ �������� ������� (����� ������� ���������� �������� ������������)

Private dimensionByX As Long                                        '����������� ������� �� ������ (x)
Private dimensionByY As Long                                        '����������� ������� �� ������ (y)

Private openSpaceLayer As Variant                               '������� ����������� �����������
Private currentgPowerLayer As Variant                        '������� �������� ������� (� ������� �� 0 �� maximumBurnPower) ����� ��������
Private gettedPowerInOneStepLayer As Variant                 '������� ���������� ������� �� ���� ��� ������� (� ����� ���� ����������� � ������� ���������)


Const cellPowerModificator As Double = 1
Const lowerBurnBound As Double = 7                          '������ ������� ��� ������� ������ �������� �������������� ������� �� �������� (� ��� �� ������ ��������������)
Const maximumBurnPower As Double = 100                      '������������ �������� ������� (����� ������� ���������� �������� ������������)

Public firedCells As Collection                             '��������� ������� ������
Public fireMaxCells As Collection                           '��������� ������� �� �������� ������ (���������� ���� �������)
Public chargedCells As Collection                           '��������� ������ ���������� � ������ ���� �������
Public startedFireCells As Collection                       '��������� ������, ������� ������ ����������, �� �� �������� ��� �� �����������
Public activeFiredCells As Collection                       '��������� �������� ������� ������ - ���, ��� ��������� � ������� - ��������� �������� ������ � ����������� ������, �� � ���������� �����������, �.�. �� ��������� �������� �������� ��� ������� ������
Public fireFrontCells As Collection                         '��������� ������ ������ ������ - ������ ������� �� ��������� � ����������� �� ������� ���� ��������������� ����
Public keyCornerCells As Collection                         '��������� �������� ������ - ������ � ������� ���� � ����� � ����� � �� �������� ������� ������
Public firePerimeterCells As Collection                     '��������� ������ ���������


'---------------------��������� �������� �������------------------------------
Public Sub CreateMatrix(ByVal dimX As Long, dimY As Long)
'���������� �������:
    dimensionByX = dimX
    dimensionByY = dimY
    
    '---���������� ���� � ������������ � ���������� �������� �������
    ReDim currentgPowerLayer(dimX, dimY) As Double
    ReDim gettedPowerInOneStepLayer(dimX, dimY) As Double
    ReDim openSpaceLayer(dimX, dimY) As Byte
    
    '---���������� ��������� ������
    Set firedCells = New Collection
    Set fireMaxCells = New Collection
    Set chargedCells = New Collection
    Set startedFireCells = New Collection
    Set activeFiredCells = New Collection
    Set fireFrontCells = New Collection
    Set keyCornerCells = New Collection
    Set firePerimeterCells = New Collection
End Sub

Public Sub SetOpenSpace(ByRef outerMatrix As Variant)
'��������� ���� ����������� �����������
    openSpaceLayer = outerMatrix
End Sub








Public Sub SetCellValue(ByVal x As Long, ByVal y As Long, ByVal layer As MatrixLayerType, ByVal value As Variant)
'������������� �������� ��� ������ � ���������� ������������ � ��������� ����
    On Error GoTo EX

    Select Case layer
        Case Is = MatrixLayerType.mtrOpenSpaceLayer
            openSpaceLayer(x, y) = value
        Case Is = MatrixLayerType.mtrCurrentgPowerLayer
            currentgPowerLayer(x, y) = value
        Case Is = MatrixLayerType.mtrGettedPowerInOneStepLayer
            gettedPowerInOneStepLayer(x, y) = value
    End Select
Exit Sub
EX:
    Debug.Print "Error: x=" & x & ", y=" & y & ", layer=" & layer & ", value=" & value
End Sub

Public Sub SetFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'������� ����� ����� �������
    
    '���� ��������� ����� ��������� �� ��������� ������� - ������� �� ���������
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then Exit Sub
    
    '���� ��������� ����� �������� ������ - ������� �� ���������
    If openSpaceLayer(x, y) = CellState.csWall Then Exit Sub
    
    On Error Resume Next
    '��������� ������ ����� �����
    If value = 0 Then
        currentgPowerLayer(x, y) = maximumBurnPower
        '��������� � ��������� ������ ������
        fireFrontCells.Add CreateNewCell(x, y), x & ":" & y
        '��������� � ��������� ������� ������
        fireMaxCells.Add CreateNewCell(x, y), x & ":" & y
    Else
        currentgPowerLayer(x, y) = value
        activeFiredCells.Add CreateNewCell(x, y), x & ":" & y
    End If
    
    '��������� ������ � ������� ����� � ���������
    firedCells.Add CreateNewCell(x, y)
End Sub

Public Sub SetFiredCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'������� ����� ������� ������ (������� ������)
    
    '���� ��������� ����� ��������� �� ��������� ������� - ������� �� ���������
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then Exit Sub
    
    '���� ��������� ����� �������� ������ - ������� �� ���������
    If openSpaceLayer(x, y) = CellState.csWall Then Exit Sub
    
    On Error Resume Next
    '��������� ������ ����� �����
    If value = 0 Then
        currentgPowerLayer(x, y) = maximumBurnPower
'        '��������� � ��������� ������ ������
'        fireFrontCells.Add CreateNewCell(x, y), x & ":" & y
        '��������� � ��������� ������� ������
        fireMaxCells.Add CreateNewCell(x, y), x & ":" & y
    Else
        currentgPowerLayer(x, y) = value
        activeFiredCells.Add CreateNewCell(x, y), x & ":" & y
    End If
    
    '���������, ��� ������ ������ �����
    openSpaceLayer(x, y) = CellState.csFire
    
    '��������� ������ � ������� ����� � ���������
    firedCells.Add CreateNewCell(x, y)
End Sub



'-----------------------------�� ��� ������� � ���������������� �������--------------------------------
Public Sub RizeCellFire(ByVal x As Long, ByVal y As Long, ByVal value As Double)
'������������������ �������� ������� ������ ������
    currentgPowerLayer(x, y) = currentgPowerLayer(x, y) + value
    If currentgPowerLayer(x, y) > maximumBurnPower Then currentgPowerLayer(x, y) = maximumBurnPower
End Sub



'-------------------------------��������--------------------------------------------
Public Property Get DimensionX() As Long
    DimensionX = dimensionByX
End Property
Public Property Get DimensionY() As Long
    DimensionY = dimensionByY
End Property
Public Property Get GetMaximumBurnPower() As Long
    GetMaximumBurnPower = maximumBurnPower
End Property


Public Function GetFirePower(ByVal x As Long, ByVal y As Long) As Double
'���������� �������� ������� ������, ��� ����, ���� ��� ���� ���������� �������� ���������������, ���������� ����
Dim tmpFirePower As Double
    
    tmpFirePower = currentgPowerLayer(x, y)
    GetFirePower = IIf(tmpFirePower < lowerBurnBound, 0, tmpFirePower)
    
End Function
Public Function GetActualFirePower(ByVal x As Long, ByVal y As Long) As Double
'� ������� �� GetFirePower ���������� �������� �������� ������� ������, ���� ���� ��� ���� ���������� �������� ���������������
    GetActualFirePower = currentgPowerLayer(x, y)
End Function


Public Function GetPowerInOneStep(ByVal x As Long, ByVal y As Long) As Double
    GetPowerInOneStep = gettedPowerInOneStepLayer(x, y)
End Function


Public Function GetOpenSpaceLayerValue(ByVal x As Long, ByVal y As Long) As Byte
    GetOpenSpaceLayerValue = openSpaceLayer(x, y)
'����� ����� ����� ������������� ������� �������� � ����������� ������� ������ - ����� ����� ������ �������� ������������� � ������ ���� ����� ���� ����� ���������:
'GetOpenSpaceLayerValue = testWall2(x * 200 / 25.4, y * 200 / 25.4, 0)
End Function
Public Function GetOpenSpaceLayer() As Variant
'���������� ������� ����������� �����������
    GetOpenSpaceLayer = openSpaceLayer
End Function
'<<<!!!!!�������� ����
Private Function testWall2(ByVal x As Double, ByVal y As Double, ByVal tolerance As Double) As Boolean
Dim shpN As Visio.Shape
Dim sel As Visio.Selection

'    Set sel = Application.ActivePage.SpatialSearch(x, y, VisSpatialRelationCodes.visSpatialOverlap, 100, visSpatialFrontToBack)
    Set sel = Application.ActivePage.SpatialSearch(x, y, VisSpatialRelationCodes.visSpatialContainedIn, 0, visSpatialFrontToBack) '- ?????????? ? ?????????? ??????
    
    For Each shpN In sel
        Debug.Print shpN.Name
        If isWallInner(shpN) Then
            testWall2 = True
            Exit Function
        End If
    Next shpN
testWall2 = False
End Function
Public Function isWallInner(ByRef aO_Shape As Visio.Shape) As Boolean
'??????? ?????????? ??????, ???? ?????? - ?????, ? ????????? ?????? - ????
Dim shapeType As Integer
'---?????????, ???????? ?? ?????? ??????? ???????????
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        isWallInner = False
        Exit Function
    End If

'---?????????, ???????? ?? ?????? ??????? ?????
    shapeType = aO_Shape.Cells("User.ShapeType").Result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (shapeType = 44 Or shapeType = 6) Then
        isWallInner = True
        Exit Function
    End If
isWallInner = False
End Function
'!!!!!�������� ����>>>

Public Function IsInner(ByVal x As Long, ByVal y As Long) As Boolean
'��������� ������, ���� ������ ������ ������ ������� (��� �����), ����, ���� ���� ���� ���� �� ��������� �������
    IsInner = True
    '�� ���������
    If IsCellCanFire(x - 1, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x - 1, y + 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y + 1) Then
        IsInner = False
        Exit Function
    End If
    '�� ����������
    If IsCellCanFire(x, y - 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x, y + 1) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x - 1, y) Then
        IsInner = False
        Exit Function
    End If
    If IsCellCanFire(x + 1, y) Then
        IsInner = False
        Exit Function
    End If
End Function

Public Function IsNotPerimeter(ByVal x As Long, ByVal y As Long) As Boolean
'��������� ������, ���� ������ ������ ������ �������, ����, ���� ���� ���� ���� �� ��������� �������
    IsNotPerimeter = False
    '�� ��������� '�� ����������
    If IsCellMaxFire(x - 1, y - 1) And _
            IsCellMaxFire(x + 1, y - 1) And _
            IsCellMaxFire(x - 1, y + 1) And _
            IsCellMaxFire(x + 1, y + 1) And _
            IsCellMaxFire(x, y - 1) And _
            IsCellMaxFire(x, y + 1) And _
            IsCellMaxFire(x - 1, y) And _
            IsCellMaxFire(x + 1, y) Then
        
        IsNotPerimeter = True
    End If
    
End Function

Public Function IsTripleStateCell(ByVal x As Long, ByVal y As Long) As Boolean
'���������� ������, ���� ������ ��������� �� ������� ���� ���� - ����, ���� � ���������� ������������ (����� ��� �� � ���������)
'��������� ��������������� � ��� ������� ������, �� ������� ������������ � �����������
Dim havingWall As Boolean
Dim havingOpenSpace As Boolean
    
    '��������� ������� ����
    If IsCellsAroundExists(x, y, 1) Then havingWall = True
    
    '��������� �� �������� �� ������ ����������
    havingOpenSpace = Not IsInner(x, y)
    
    IsTripleStateCell = havingWall And havingOpenSpace

End Function


Public Function IsCellsAroundExists(ByVal x As Long, ByVal y As Long, ByVal value As Integer) As Boolean
'���������� ������, ���� ������ ������, � ���� ����������� ����������� ����� ��������� �������� (���� ����), �����, ����
    '�� ���������
    If IsCellsAreType(x - 1, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x - 1, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    '�� ����������
    If IsCellsAreType(x, y - 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x, y + 1, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x - 1, y, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
    If IsCellsAreType(x + 1, y, value) Then
        IsCellsAroundExists = True
        Exit Function
    End If
IsCellsAroundExists = False
End Function
Private Function IsCellsAreType(ByVal x As Long, ByVal y As Long, ByVal value As Integer) As Boolean
'��������� ������ ���� ������ ������ (��� ���� ����������� �����������)
    '���������, ������ �� ������ � ������� �������
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
        IsCellsAreType = False
        Exit Function
    End If
    
    IsCellsAreType = openSpaceLayer(x, y) = value
End Function

'Private Function IsCellsAroundNotMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
''���������, ������� �� � ��������� ������ �� ������� �� ������ ��������
'    '�� ���������
'    If IsCellNotMaxFire(x - 1, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x - 1, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    '�� ����������
'    If IsCellNotMaxFire(x, y - 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x, y + 1) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x - 1, y) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'    If IsCellNotMaxFire(x + 1, y) Then
'        IsCellsAroundNotMaxFire = True
'        Exit Function
'    End If
'IsCellsAroundNotMaxFire = False
'End Function
'Private Function IsCellNotMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
''���������� ������, ���� ������ ����� �� �� ������ �������� (���� ������ �� �����) � ����, ���� ���
'    '���������, ������ �� ������ � ������� �������
'    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
'        IsCellNotMaxFire = False
'        Exit Function
'    End If
'
'    IsCellNotMaxFire = currentgPowerLayer(x, y) < maximumBurnPower
'End Function

Public Function IsCellOutOfArea(ByVal x As Long, ByVal y As Long) As Boolean
'���������, ������� �� ������ �� ������� ������� �������
    If TypeName(openSpaceLayer(x, y)) = "Empty" Then
        IsCellOutOfArea = True
        Exit Function
    End If
IsCellOutOfArea = False
End Function

Private Function IsCellCanFire(ByVal x As Long, ByVal y As Long) As Boolean
'���������� ������, ���� ������ ����� ������ � ����, ���� ���
    
    '���������, ������ �� ������ � ������� �������
'    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
'        IsCellCanFire = False
'        Exit Function
'    End If
    If IsCellInMatrix(x, y) = False Then
        IsCellCanFire = False
        Exit Function
    End If
    
    '���������, �� �������� �� ������ ������
    If openSpaceLayer(x, y) = CellState.csWall Then
        IsCellCanFire = False
        Exit Function
    End If
    
    '��������� �������� ������� ������
    If currentgPowerLayer(x, y) < maximumBurnPower Then
        IsCellCanFire = True
    Else
        IsCellCanFire = False
    End If
    
End Function

Public Function IsCellInMatrix(ByVal x As Long, ByVal y As Long) As Boolean
'�������� �� ���������� ������ � ������ ������������ � �������
    If x < 0 Or y < 0 Or x > DimensionX Or y > DimensionY Then
        IsCellInMatrix = False
        Exit Function
    End If
IsCellInMatrix = True
End Function


Public Function isWall(ByVal x As Long, ByVal y As Long) As Boolean
'���������, ��������� �� � ��������� ������ �����
    On Error GoTo EX
    If openSpaceLayer(x, y) = 1 Then
        isWall = True
    Else
        isWall = False
    End If
Exit Function
EX:
    isWall = True
End Function

Public Function IsCellMaxFire(ByVal x As Long, ByVal y As Long) As Boolean
'���������, ����� �� ������ �� ������ ��������
    On Error GoTo EX
    IsCellMaxFire = currentgPowerLayer(x, y) = maximumBurnPower
Exit Function
EX:
    IsCellMaxFire = False
End Function

Public Function IsOuterActiveCell(ByVal x As Long, ByVal y As Long) As Boolean
'���������, �������� �� ������ ������� ��������, �.�. ��� �� ����� �� ������ ������� � ��������� �� ��������� ������� �������
Dim val As Double
    
    On Error GoTo EX
    
    val = GetFirePower(x, y)
    IsOuterActiveCell = val > 0 And val < maximumBurnPower
        
Exit Function
EX:
    
End Function

'-----------------------------�����������-------------------------------------------
Private Function CreateNewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'������� ����� ������
Dim cell As c_Cell
    
    Set cell = New c_Cell
    Set cell.ParentMatrix = Me
    cell.x = x
    cell.y = y
    
Set CreateNewCell = cell
End Function
