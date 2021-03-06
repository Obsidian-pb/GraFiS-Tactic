VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FrontDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim matrix As Variant
Private dimensionByX As Long                                        '����������� ������� �� ������ (x)
Private dimensionByY As Long                                        '����������� ������� �� ������ (y)

Public grain As Integer



'---------------------------����� ���������� �� ������������ ��������� ������ �������---------------------------

Private Enum CellType
    notFireFront = 0
    fireFrontNotInBlock = 1
    fireFrontInBlock = 2
End Enum

Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'���������� ������� � ����������� � ����������� ���������
    dimensionByX = x
    dimensionByY = y
    ReDim matrix(x, y)
End Sub

Public Sub SetStartState(ByRef frontCellsCollection As Collection)
'������������� ��������� ��������� �������
    '1 ������ ��������� ��������� ������� 5 - ����
    For Each Cell In frontCellsCollection
        matrix(Cell.x, Cell.y) = CellType.fireFrontNotInBlock
    Next Cell
End Sub




Public Sub RunDemon(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection, _
                    ByRef nozzle As c_FireNozzleModel)
'��������� ������� ������ �� �����
Dim block As c_ExtinguishingBlock

    '1 ���������, �� ������ �� ����� �������� ���� ������ ������ � �������� ������ ������, ������� ��� �����������
    Set block = CheckBlockExists(extSquare, nozzle.fireContactCell)
    If block Is Nothing Then
        '2 ������� ��������� ������ �� ������� �� ��������� ����� �����
        BreakFront extSquare, frontCellsCollection, nozzle
    Else
        '3 ��������� ������ � ������������ ����
        block.expenseFact = block.expenseFact + nozzle.waterExpense
    End If
    


End Sub




Private Sub BreakFront(ByRef extSquare As c_ExtinguishingSquare, ByRef frontCellsCollection As Collection, _
                    ByRef nozzle As c_FireNozzleModel)
'��������� ����� ������ �� �������
'Dim cell As c_Cell
Dim extBlock As c_ExtinguishingBlock
Dim tmpColl As Collection

    '���������� ��� ������ � ��������� ��������� ������
'    For Each cell In startCellsCollection
        
        '���� ������ ��� �� ������������ �� ������, �������� �� ��� ����������� ����� ���� ������ ������� ������� ������� �����
'        If matrix(startCell.x, startCell.y) = CellType.fireFrontNotInBlock Then
            '������� ����� ���������
            Set tmpColl = New Collection
            '��� ����� ��������� ��������� ����������� ���������� �� ������� ������
            TryAddCell tmpColl, nozzle.fireContactCell
            
            '������� ����� ���� � �������� ��� ��������� ��������� ������ ������
            Set extBlock = New c_ExtinguishingBlock
            Set extBlock.fireFrontCells = tmpColl
            extBlock.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
            extBlock.AddToExtCellsCollection tmpColl
'            AddCellsToCollection extBlock.extinguishingCells, tmpColl
            '---������������� ������� �������
                If nozzle.nozzleType = waterHand Then
                    extBlock.fireWay = 5000     '��� ������ �������
                Else
                    extBlock.fireWay = 10000     '��� �������� �������
                End If
                extBlock.expenseFact = extBlock.expenseFact + nozzle.waterExpense
                extBlock.grain = grain
            
            '��������� ��������� ���� � ��������� ������ ������ ������� �������
            extSquare.AddExtinguishingBlock extBlock
'        End If
'    Next cell
        
End Sub

Private Sub TryAddCell(ByRef thisColl As Collection, ByVal Cell As c_Cell)
'�������� �������� ������ � ��������� (��������)
    
    '���� ��������� ����� ��������� �� ��������� ������� - ������� �� ���������
    If Cell.x < 0 Or Cell.y < 0 Or Cell.x > dimensionByX - 1 Or Cell.y > dimensionByY - 1 Then Exit Sub
    
    '��������� �������� �� ������ ������ ������� ������ ������
    If matrix(Cell.x, Cell.y) = CellType.notFireFront Then Exit Sub
    
    '���� �������� ���������, ��� �� �� ��� � ���������
    If matrix(Cell.x, Cell.y) = CellType.fireFrontInBlock Then Exit Sub
    
    '��������� ������ � ��������
    AddCellToCollection thisColl, Cell
    matrix(Cell.x, Cell.y) = CellType.fireFrontInBlock
    
    '���� ��� - �������� �� ��������� ������
        TryAddCell thisColl, NextCell(Cell, l)
        TryAddCell thisColl, NextCell(Cell, lu)
        TryAddCell thisColl, NextCell(Cell, u)
        TryAddCell thisColl, NextCell(Cell, ru)
        TryAddCell thisColl, NextCell(Cell, r)
        TryAddCell thisColl, NextCell(Cell, rd)
        TryAddCell thisColl, NextCell(Cell, d)
        TryAddCell thisColl, NextCell(Cell, ld)
   
End Sub

Private Sub AddCellToCollection(ByRef thisColl As Collection, ByVal Cell As c_Cell)
'��������� � ��������� ��������� ��������� ������
    thisColl.Add Cell, Cell.x & ":" & Cell.y
End Sub

Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'��������� ��� ������ �� ��������� 2 � ��������� 1
Dim Cell As c_Cell

    For Each Cell In col2
        col1.Add Cell, Cell.x & ":" & Cell.y
    Next Cell
End Sub

Private Function CheckBlockExists(ByRef extSquare As c_ExtinguishingSquare, ByRef Cell As c_Cell) As c_ExtinguishingBlock
'�������� ����� ���� � �������� ����� ���������� ��������� ������. ������ ������ ���������� ������ �� ����, ����� - Nothing
Dim blck As c_ExtinguishingBlock

    For Each blck In extSquare.extinguishingBlocks
        If blck.IsCellInThisFront(Cell) Then
            Set CheckBlockExists = blck
            Exit Function
        End If
    Next blck
Set CheckBlockExists = Nothing
End Function

'Public Function DrawBlockExists(ByRef extSquare As c_ExtinguishingSquare, ByRef cell As c_Cell) As c_ExtinguishingBlock
''�������� ����� ���� � �������� ����� ���������� ��������� ������. ������ ������ ���������� ������ �� ����, ����� - Nothing
'Dim blck As c_ExtinguishingBlock
'
'Dim c As c_Cell
'Dim shp As Visio.Shape
'Const mmInInch As Double = 25.4
'
'    For Each blck In extSquare.extinguishingBlocks
'        For Each c In blck.fireFrontCells
'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle(c.x * grain / mmInInch, c.y * grain / mmInInch, _
'                    (c.x + 1) * grain / mmInInch, (c.y + 1) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Front")
'            shp.Cells("LinePattern").FormulaU = "0"
'            shp.Cells("FillForegnd").FormulaU = "3"
'        Next c
'
'
''        If blck.IsCellInThisFront(cell) Then
''            Set CheckBlockExists = blck
''            Exit Function
''        End If
'    Next blck
'Set CheckBlockExists = Nothing
'End Function



Private Function NextCell(ByRef Cell As c_Cell, ByVal dir As Directions) As c_Cell
'���������� ��������� ������ �������� ���������� �����������
    Select Case dir
        Case Is = Directions.l
            Set NextCell = NewCell(Cell.x - 1, Cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(Cell.x - 1, Cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(Cell.x, Cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(Cell.x + 1, Cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(Cell.x + 1, Cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(Cell.x + 1, Cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(Cell.x, Cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(Cell.x - 1, Cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'������� ����� ������
Dim Cell As c_Cell
    Set Cell = New c_Cell
    Cell.x = x
    Cell.y = y
    
    Set NewCell = Cell
End Function

Public Property Get GetMatrix() As Variant
    GetMatrix = matrix
End Property
