VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtinguishingSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public extinguishingBlocks As Collection           'Коллекция блоков тушения
Public intenseNeed As Single                       'Требуемая интенсивность подачи воды л/(с*м.кв.)


'--------------------------------Класс зоны площади тушения (модель)-------------------------------------------

Private Sub Class_Initialize()
    Set extinguishingBlocks = New Collection
End Sub

Public Sub AddExtinguishingBlock(ByRef block As c_ExtinguishingBlock)

    block.LetID = extinguishingBlocks.Count
    block.intenseNeed = intenseNeed
    extinguishingBlocks.Add block

End Sub

Public Function IsCellExtinguished(ByVal x As Long, ByVal y As Long) As WaterExpenseKind
'Проверяем имеется ли входящая клетка в коллекциях тушащихся в каждом из блоков
Dim block As c_ExtinguishingBlock
    
    For Each block In extinguishingBlocks
        If block.IsCellExtinguished(x, y) Then
            IsCellExtinguished = block.GetWaterExpenseKind

            Exit Function
        End If
    Next block
    
IsCellExtinguished = notSet
End Function

Public Function GetSquare() As Single
'Функция возвращает площадь тушения всех блоков, т.е. общую площадь тушения для данного пожара
Dim block As c_ExtinguishingBlock
Dim tmpSquare As Single
    
    For Each block In extinguishingBlocks
        tmpSquare = tmpSquare + block.GetSquare
    Next block
GetSquare = tmpSquare
End Function

Public Sub DeleteCellFromEveryBlock(ByVal x As Long, ByVal y As Long)
'Удаляяем клетку из всех блоков площади тушения
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
'Удаляяем клетку из всех коллекций TmpCells блоков площади тушения
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
'Подается воды: 0 не подается, 1 - подается недостаточно, 2 - подается достаточно.
Dim block As c_ExtinguishingBlock
Dim expKind(2) As Byte
    
    'Если нет ни одного блока тушения возвращаем значение notSet
    If extinguishingBlocks.Count = 0 Then
        GetWaterExpenseKind = notSet
        Exit Function
    End If
    
    'Проверяем все блоки - достаточно ли расхода
    For Each block In extinguishingBlocks
        expKind(block.GetWaterExpenseKind) = expKind(block.GetWaterExpenseKind) + 1
    Next block
    
    'В зваисимости от того, что получилось укзываем достаточность расхода для всей площади тушения
    If expKind(sufficient) = extinguishingBlocks.Count Then
        GetWaterExpenseKind = sufficient
        Exit Function
    End If
    If expKind(notSet) = extinguishingBlocks.Count Then
        GetWaterExpenseKind = notSet
        Exit Function
    End If
    
GetWaterExpenseKind = notSufficient
End Function
