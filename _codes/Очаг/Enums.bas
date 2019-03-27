Attribute VB_Name = "Enums"
'Перечисление состояния клетки (возможно, не нужно)
Enum CellState
    csOpenSpace = 0
    csWall = 1
    csFire = 2
    csFireOuter = 3
    csFireInner = 4
    csWillBurnNextStep = 5

End Enum

'Перечисление возможных типов слоев в матрице
Enum MatrixLayerType
    mtrOpenSpaceLayer = 0
    mtrCurrentgPowerLayer = 1
    mtrGettedPowerInOneStepLayer = 2
End Enum

'Перечисление возможных направлений движения клеточных демонов
Enum Directions
    s = 0       'Стоит
    l = 1       'Влево
    lu = 2      'Влево вверх
    u = 3       'Вверх
    ru = 4      'Вправо вверх
    r = 5       'Вправо
    rd = 6      'Вправо вниз
    d = 7       'Вниз
    ld = 8      'Влево вниз
End Enum

'Типы пожарных стволов
Enum NozzleTypes
    waterHand = 0
    waterLafet = 1
    foamHand = 2
    foamLafet = 3
    powderHand = 4
    powderLafet = 5
    gas = 6
    
End Enum

'Тип достаточности подаваемой воды
Enum WaterExpenseKind
    notSet = 0
    notSufficient = 1
    sufficient = 2
End Enum

'Наличие ячейки в зоне тушения
Enum ExtCellType
    notInExtSquare = 0
    inExtSquare = 1
End Enum


'Enum CellSpreadType
'    csstInner = 0
'    csstSingleton = 1
'    csstCannon = 2
'    csstHardCannon = 3
'
'End Enum
