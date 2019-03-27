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

'Public Function IsCellExtinguished(ByVal x As Long, ByVal y As Long) As Boolean
''Проверяем имеется ли входящая клетка в коллекциях тушащихся в каждом из блоков
'Dim block As c_ExtinguishingBlock
'
'    For Each block In extinguishingBlocks
'        If block.IsCellExtinguished(x, y) Then
'            IsCellExtinguished = True
'            Exit Function
'        End If
'    Next block
'
'IsCellExtinguished = False
'End Function
