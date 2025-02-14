VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Integer                           'ID ����� - ����� ��� �������������, ���������� � 1.

Private cellsMatrix As Variant                  '������� ������

Public extinguishingCells As Collection         '��������� ������ ���� ������� �������, ����������� � ������� �����
Public fireFrontCells As Collection             '��������� ������ ������ ������, ����������� � �������� �����
Public activeCells As Collection                '��������� ������ �������� ���� (��� �� ������� �� ��������), ����������� � �������� �����
Public tmpColl As Collection                    '��������� ��������� ������ - ��� ������� ������� �������

Public fireWay As Single                        '������� ������� ��������� �������� (� �����������)
Public grain As Long
Public intenseNeed As Single                    '��������� ������������� ������ ���� �/(�*�.��.)
Public expenseFact As Single                    '����� ���������� ���� �/�

Public alreadyCalculated As Boolean             '����, ��� ���� ��� ��������� (��������� ���������)

'--------------------����� ����� ���� ������� (������)---------------------------------------------------



Private Sub Class_Initialize()
    Set extinguishingCells = New Collection
    Set fireFrontCells = New Collection
    Set activeCells = New Collection
    Set tmpColl = New Collection
    
    alreadyCalculated = False
End Sub

'-------------------------------������ � ��������---------------------------------
Public Sub CreateMatrix(ByVal x As Integer, ByVal y As Integer)
'���������� ������� � ����������� � ����������� ���������
    ReDim cellsMatrix(x, y)
End Sub

Public Sub AddToExtCellsCollection(ByRef coll As Collection)
'��������� ��������� �������� ������ � ��������� �������
Dim cell As c_Cell
    
    AddCellsToCollection extinguishingCells, coll
    
    '��� ���� �������� ������ ������������� �������� "� ���� �������"
    For Each cell In coll
        cellsMatrix(cell.x, cell.y) = ExtCellType.inExtSquare
    Next cell
End Sub
Public Sub RemoveFromExtCellsCollection(ByRef cell As c_Cell)
'��������� �������� ������ �� ��������� ���������
    
    On Error Resume Next
    
    extinguishingCells.Remove cell.x & ":" & cell.y
    cellsMatrix(cell.x, cell.y) = ExtCellType.notInExtSquare

End Sub
Public Sub RemoveFromTmpCellsCollection(ByRef cell As c_Cell)
'��������� �������� ������ �� ��������� ��������� ��� �������
    
    On Error Resume Next
    
    tmpColl.Remove cell.x & ":" & cell.y
'    cellsMatrix(cell.x, cell.y) = ExtCellType.notInExtSquare

End Sub

Public Sub AddToActiveCellsCollection(ByRef coll As Collection)
'��������� ��������� �������� ������ � ��������� �������� ������ (�������, �������������)
Dim cell As c_Cell
    
    AddCellsToCollection activeCells, coll
    
    '��� ���� �������� ������ ������������� �������� "� ���� �������"
    On Error Resume Next
    For Each cell In coll
        cellsMatrix(cell.x, cell.y) = ExtCellType.inExtSquare
    Next cell
End Sub

Public Function IsCellExtinguished(ByVal x As Long, ByVal y As Long) As Boolean
    IsCellExtinguished = cellsMatrix(x, y) = ExtCellType.inExtSquare
End Function



Public Function GetSquare() As Single
'���������� ������� ���� ������� � ������ ����������
    GetSquare = (extinguishingCells.Count * grain ^ 2) / 1000000
End Function

Public Function GetExpenceNeed() As Single
'���������� ��������� ������ ���� ��� ������� ������ �������
    GetExpenceNeed = GetSquare * intenseNeed
End Function

Public Function GetWaterExpenseKind() As WaterExpenseKind
'�������� ����: 0 �� ��������, 1 - �������� ������������, 2 - �������� ����������.
Dim expenseNeed As Single
    
    expenseNeed = GetExpenceNeed
    If expenseFact = 0 Then
        GetWaterExpenseKind = notSet
    End If
    If expenseFact < expenseNeed Then
        GetWaterExpenseKind = notSufficient
    End If
    If expenseFact >= expenseNeed Then
        GetWaterExpenseKind = sufficient
    End If
    
End Function


Public Function IsCellInThisFront(ByRef cell As c_Cell) As Boolean
'���������, ��������� �� ��������� ������ � ��������� ������ ������ ������ ������� �����
Dim cll As c_Cell
    For Each cll In fireFrontCells
        If cll.x = cell.x And cll.y = cell.y Then
            IsCellInThisFront = True
            Exit Function
        End If
    Next
IsCellInThisFront = False
End Function

Public Function IsCellInThisPerimeter(ByRef cell As c_Cell) As Boolean
'���������, ��������� �� ��������� ������ � ��������� ������ ��������� ������ ������� �����
Dim cll As c_Cell
    
    IsCellInThisPerimeter = Not (IsCellInThisBlock(cell.NCell(d)) And _
                            IsCellInThisBlock(cell.NCell(l)) And _
                            IsCellInThisBlock(cell.NCell(ld)) And _
                            IsCellInThisBlock(cell.NCell(lu)) And _
                            IsCellInThisBlock(cell.NCell(r)) And _
                            IsCellInThisBlock(cell.NCell(rd)) And _
                            IsCellInThisBlock(cell.NCell(ru)) And _
                            IsCellInThisBlock(cell.NCell(u)))
End Function

Public Function IsCellInThisBlock(ByRef cell As c_Cell) As Boolean
'���������, ��������� �� ��������� ������ � ��������� ������ ������� �����
Dim tmpCell As c_Cell
    On Error GoTo EX
    '�������� �������� �� ����� �� �������!!!
    
    Set tmpCell = extinguishingCells.item(cell.x & ":" & cell.y)
    IsCellInThisBlock = True
Exit Function
EX:
'    Debug.Print cell.x * grain & ":" & cell.y * grain
    IsCellInThisBlock = False
End Function


'--------------------------��������---------------------------------------------------
Property Let LetID(ByVal val As Integer)
    ID = val
End Property
Property Get GetID()
    GetID = ID
End Property




'------------------------�����������-------------------------------------------------
Private Sub AddCellsToCollection(ByRef col1 As Collection, ByRef col2 As Collection)
'��������� ��� ������ �� ��������� 2 � ��������� 1
Dim cell As c_Cell
    
    On Error Resume Next
    For Each cell In col2
        col1.Add cell, cell.x & ":" & cell.y
    Next cell
End Sub
