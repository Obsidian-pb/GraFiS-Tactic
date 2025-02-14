VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public extinguishingBlocks As Collection           '��������� ������ �������
Public intenseNeed As Single                       '��������� ������������� ������ ���� �/(�*�.��.)


'--------------------------------����� ���� ������� ������� (������)-------------------------------------------

Private Sub Class_Initialize()
    Set extinguishingBlocks = New Collection
End Sub

Public Sub AddExtinguishingBlock(ByRef block As c_ExtinguishingBlock)

    block.LetID = extinguishingBlocks.Count
    block.intenseNeed = intenseNeed
    extinguishingBlocks.Add block

End Sub

Public Function IsCellExtinguished(ByVal x As Long, ByVal y As Long) As WaterExpenseKind
'��������� ������� �� �������� ������ � ���������� ��������� � ������ �� ������
Dim block As c_ExtinguishingBlock
    
    For Each block In extinguishingBlocks
        If block.IsCellExtinguished(x, y) Then
            IsCellExtinguished = block.GetWaterExpenseKind

            Exit Function
        End If
    Next block
    
    
IsCellExtinguished = notSet
End Function

Public Function IsSquareExtinguished(ByRef nozzleDemon As c_FireNozzlesDemon, ByVal curTime As Single) As WaterExpenseKind
'���������� ������������� �������
Dim ExpNeed As Double
Dim ExpFact As Double

    ExpNeed = GetSquare * intenseNeed
    ExpFact = nozzleDemon.GetExpense(curTime)
        
    If ExpFact = 0 Then
        IsSquareExtinguished = notSet
        Exit Function
    End If
    
    If ExpNeed < ExpFact Then
        IsSquareExtinguished = sufficient
    Else
        IsSquareExtinguished = notSufficient
    End If
    
End Function

Public Function GetSquare() As Single
'������� ���������� ������� ������� ���� ������, �.�. ����� ������� ������� ��� ������� ������
Dim block As c_ExtinguishingBlock
Dim tmpSquare As Single
    
    For Each block In extinguishingBlocks
        tmpSquare = tmpSquare + block.GetSquare
    Next block
GetSquare = tmpSquare
End Function

Public Sub DeleteCellFromEveryBlock(ByVal x As Long, ByVal y As Long)
'�������� ������ �� ���� ������ ������� �������
Dim cell As c_Cell
Dim block As c_ExtinguishingBlock
    
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    For Each block In extinguishingBlocks
        block.RemoveFromExtCellsCollection cell
    Next block
End Sub
Public Sub DeleteCellTmpFromEveryBlock(ByVal x As Long, ByVal y As Long)
'�������� ������ �� ���� ��������� TmpCells ������ ������� �������
Dim cell As c_Cell
Dim block As c_ExtinguishingBlock
    
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    For Each block In extinguishingBlocks
        block.RemoveFromTmpCellsCollection cell
    Next block
End Sub

Public Function IsAllBlocksCalculated() As Boolean
Dim block As c_ExtinguishingBlock
    
    For Each block In extinguishingBlocks
        If block.alreadyCalculated = False Then
            IsAllBlocksCalculated = False
            Exit Function
        End If
    Next block
IsAllBlocksCalculated = True
End Function

Public Function GetWaterExpenseKind() As WaterExpenseKind
'�������� ����: 0 �� ��������, 1 - �������� ������������, 2 - �������� ����������.
Dim block As c_ExtinguishingBlock
'Dim expKind(2) As Byte
Dim expKind As WaterExpenseKind
    
    '������������� ��������� �������� ������������� ������� - �� ���������, ����������
    expKind = WaterExpenseKind.sufficient
    
    '���� ��� �� ������ ����� ������� ���������� �������� notSet
    If extinguishingBlocks.Count = 0 Then
        GetWaterExpenseKind = notSet
        Exit Function
    End If
    
'    '��������� ��� ����� - ���������� �� �������
'    For Each block In extinguishingBlocks
'        expKind(block.GetWaterExpenseKind) = expKind(block.GetWaterExpenseKind) + 1
'    Next block
'
'    '� ����������� �� ����, ��� ���������� �������� ������������� ������� ��� ���� ������� �������
'    If expKind(sufficient) = extinguishingBlocks.Count Then
'        GetWaterExpenseKind = sufficient
'        Exit Function
'    End If
'    If expKind(notSet) = extinguishingBlocks.Count Then
'        GetWaterExpenseKind = notSet
'        Exit Function
'    End If
    
    '��������� ��� ����� - ���������� �� �������
    For Each block In extinguishingBlocks
        If expKind > block.GetWaterExpenseKind Then
            expKind = block.GetWaterExpenseKind
        End If
    Next block

'GetWaterExpenseKind = notSufficient
GetWaterExpenseKind = expKind
End Function
