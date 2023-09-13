Attribute VB_Name = "Enums"
'������������ ��������� ������ (��������, �� �����)
Enum CellState
    csOpenSpace = 0
    csWall = 1
    csFire = 2
    csFireOuter = 3
    csFireInner = 4
    csWillBurnNextStep = 5

End Enum

'������������ ��������� ����� ����� � �������
Enum MatrixLayerType
    mtrOpenSpaceLayer = 0
    mtrCurrentgPowerLayer = 1
    mtrGettedPowerInOneStepLayer = 2
End Enum

'������������ ��������� ����������� �������� ��������� �������
Enum Directions
    s = 0       '�����
    l = 1       '�����
    lu = 2      '����� �����
    u = 3       '�����
    ru = 4      '������ �����
    r = 5       '������
    rd = 6      '������ ����
    d = 7       '����
    ld = 8      '����� ����
End Enum

'���� �������� �������
Enum NozzleTypes
    waterHand = 0
    waterLafet = 1
    foamHand = 2
    foamLafet = 3
    powderHand = 4
    powderLafet = 5
    gas = 6
    
End Enum

'��� ������������� ���������� ����
Enum WaterExpenseKind
    notSet = 0
    notSufficient = 1
    sufficient = 2
End Enum

'������� ������ � ���� �������
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
