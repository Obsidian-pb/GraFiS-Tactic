VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Modeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------����� ������������ ��������� ������----------------------------------------------
Const diag As Double = 0.1                                  ' ����������� �������� �� ���������
Const orto As Double = 0.14142135623731                     ' ����������� �������� �� ��������� � �����������
Const mmInInch As Double = 25.4

Private matrix As c_Matrix
Private extSquare As c_ExtinguishingSquare                  '������ ������� �������

Public grain As Double              '������ ����� ������� (������)
Private step As Integer             '������� ���������� ����� (�����, ��� ������������� � � ��������������)
Private calcStep As Integer         '������� ���������� ������������� (������ � ��������������) �����
Public distance As Single           '����� ��������� ���������� ��� ������������� (�������� ��� ���������� ������)
Public time As Single               '����� ����� ��������� ��� ������������� (�������� ��� ���������� ������)
Public intenseNeed As Single        '��������� ������������� ������ ���� �/(�*�.��.)

Private nozzleDemon As c_FireNozzlesDemon
    


Public Sub SetMatrix(ByRef m As c_Matrix)
'������������� ������ �� �������
    Set matrix = m
End Sub

Public Sub refreshOpenSpaces(ByRef outerMatrix() As Variant)
'��������� ������� �������� �����������
    matrix.SetOpenSpace outerMatrix
End Sub

Public Sub SetFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'������� ����� ����� �������
    matrix.SetFireCell x, y, value
    
    '��������� ������ � ������� ����� � ��������� �������� ������
    AddToActiveCells x, y
    
End Sub

Public Sub SetStartFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'������� ��������� ����� �������
    matrix.SetFireCell x, y, value
    '��������� ����� ��������������� �������
    matrix.SetFireCell x - 2, y - 2, 3.4
    matrix.SetFireCell x - 1, y - 2, 9.61
    matrix.SetFireCell x, y - 2, 13.59
    matrix.SetFireCell x + 1, y - 2, 9.61
    matrix.SetFireCell x + 2, y - 2, 3.4

    matrix.SetFireCell x - 2, y - 1, 9.61
    matrix.SetFireCell x - 1, y - 1, 43.58
    matrix.SetFireCell x, y - 1, 61.64
    matrix.SetFireCell x + 1, y - 1, 43.58
    matrix.SetFireCell x + 2, y - 1, 9.61

    matrix.SetFireCell x - 2, y, 13.59
    matrix.SetFireCell x - 1, y, 61.64
    matrix.SetFireCell x + 1, y, 61.64
    matrix.SetFireCell x + 2, y, 13.59

    matrix.SetFireCell x - 2, y + 1, 9.61
    matrix.SetFireCell x - 1, y + 1, 43.58
    matrix.SetFireCell x, y + 1, 61.64
    matrix.SetFireCell x + 1, y + 1, 43.58
    matrix.SetFireCell x + 2, y + 1, 9.61

    matrix.SetFireCell x - 2, y + 2, 3.4
    matrix.SetFireCell x - 1, y + 2, 9.61
    matrix.SetFireCell x, y + 2, 13.59
    matrix.SetFireCell x + 1, y + 2, 9.61
    matrix.SetFireCell x + 2, y + 2, 3.4
    
    '��������� ������ � ������� ����� � ��������� �������� ������
    AddToActiveCells x, y
    
End Sub

Public Function AddFireFromShape(ByRef fireShape As Visio.Shape) As c_Modeller
'���������� � ������� ������ ��������� ������
Dim fireShapeDemon As c_FireShapeDemon

    Set fireShapeDemon = New c_FireShapeDemon
    '����������� ������� ������� ������
    fireShapeDemon.setGrain(grain).SetFireShape(fireShape).RunDemon matrix

    Set fireShapeDemon = Nothing
    
Set AddFireFromShape = Me
End Function

Public Sub ActivateNozzles(ByRef frm As Object)
    Set nozzleDemon = New c_FireNozzlesDemon
    nozzleDemon.SetForm(frm).setGrain = grain
    Set nozzleDemon.SetMatrix = matrix
    nozzleDemon.FindNozzles
End Sub



Public Sub OneRound()
'���� ��� �������������
Dim cell As c_Cell
    
    '����������� � ���� - ��������� ��������������� ���� ���� ��� �������� ������!!!
  
    
    '������� ��������� ������ ������ ���������� ������ (�������������)
    Set matrix.startedFireCells = New Collection
    
    '������� ��������� �������� �����
    Set matrix.keyCornerCells = New Collection

    '������� ����� ��������� �������� (�� �����������)
    For Each cell In matrix.activeFiredCells
        Attack cell.x, cell.y
    Next cell
    
    '--------------------
    '����� ����� ����������� �������� ������� ���� ������ � ������������ � ���������� ��� ���������
    For Each cell In matrix.chargedCells
        RizeCellFire cell.x, cell.y
    Next cell
    
    '������� ��������� ������ ���������� �����
    Set matrix.chargedCells = New Collection
    
    '���� �� ������� - ���� ���� ����� �� ������� ���...
'    '��������� ������, �� ������� ����, �� ����� �� ��� �����������, � � ���� ������ ������� �� �� ���������
'    For Each cell In matrix.activeFiredCells
'        If matrix.IsInner(cell.x, cell.y) Then
'            RemoveFromActiveCells cell.x, cell.y
'            RemoveFromFrontCells cell.x, cell.y
'        End If
'        If matrix.IsNotPerimeter(cell.x, cell.y) Then
'            RemoveFromPerimeterCells cell.x, cell.y
'        End If
'    Next cell

    
    '��������� ������� � ���� (��������� ��� ������, ��� �������� �������, �� ��� �� ����� ���������))
    If F_InsertFire.Flag_CornerDemonIsActive.value = True Then
        Dim shp As Visio.Shape
        Dim cornerDemon As c_CornerDemon
        Set cornerDemon = New c_CornerDemon
        cornerDemon.Init matrix, 0, 0
        For Each cell In matrix.fireFrontCells
            If matrix.IsTripleStateCell(cell.x, cell.y) Then
            
            '!!!������������ ������� ���� ������ �������� ��������
    
    '            Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
    '                    (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
    '            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������� �����")
    '            shp.Cells("LinePattern").FormulaU = "0"
    '            shp.Cells("FillForegnd").FormulaU = "3"
                
                '��������� ������ �����������
                cornerDemon.SetStartPosition cell.x, cell.y
                cornerDemon.SetDirection s
                cornerDemon.RunDemon
    
                '��������� ������� ����� ���� - ��� ���������� ��������� ��������, ����� ����� ����� ��������� � ���� ������������ (��������� ��� ��������� ����)
                cornerDemon.SetDirection s
                cornerDemon.RunDemon
            End If
        Next cell

    '    Debug.Print "����� ��� �������" & cornerDemon.lineFillDemon.cellsForFill.Count
        For Each cell In cornerDemon.lineFillDemon.cellsForFill
            RizeCellFireToMax cell.x, cell.y
        Next cell
    End If
    
    
    '��������� ������, �� ������� ����, �� ����� �� ��� �����������, � � ���� ������ ������� �� �� ���������
    For Each cell In matrix.activeFiredCells
        If matrix.IsInner(cell.x, cell.y) Then
            RemoveFromActiveCells cell.x, cell.y
            RemoveFromFrontCells cell.x, cell.y
        End If
        If matrix.IsNotPerimeter(cell.x, cell.y) Then
            RemoveFromPerimeterCells cell.x, cell.y
        End If
    Next cell

    '������������ ������� ������� ��� ������� ���������� �� �������� ������, ��� �������, ��� ��� ���� ������ ����� �������� � ������� ������
    '��������� ������ ������� ��������� �������
    NozllesRecalculate


    '����������� ����� ������������ ����
    calcStep = calcStep + 1
End Sub


Private Sub Attack(ByVal x As Long, ByVal y As Long)
'������� ��������� ������� ��� �������� � ����������� � ��������� ������� � ������ ���������������
Dim firePower As Double
Dim delta1Stvol As Date
    
    '0 ��������� ��������� �� ������ �� �����
    If matrix.IsCellInMatrix(x, y) = False Then
        Exit Sub
    End If
    
    '1 �������� �������� ������� ������
    firePower = matrix.GetFirePower(x, y)
    
    '2 ��������� �������������� �� ������ ������� � �������� �� ��� �������� (�����)
    If firePower = 0 Or matrix.IsInner(x, y) Then
        Exit Sub
    End If
    
    '3 ��������� ����� ������� ������� �������� ������� ������
'��� ��� ������� ������������ � ����������� ������� ������� (���� ������� ����� ������ ������ �����, �� ������� ������� ������� ��������� �� ����� ������ ������)
'    If F_InsertFire.CB_StvolsOptions.Text = "������� ��������� ������" Then
'        '3.1 ��������� ��������� �� ������� ��������� ����� � ��������� 10 ����� - ����� ������ 1 ������ (��1) � ������ �� ����� ������� ���������������
'        ' ��� ������ ������� ����� ������ ���� ������ 10 � ������ ������ ������ ������� ������ (� ������� ������ ������)
'        delta1Stvol = DateDiff("n", F_InsertFire.VmD_TimeStart, F_InsertFire.VmD_Time1Stvol)
'        If time > 10 And time < delta1Stvol And isOdd(step) Then
'            Exit Sub
'        End If
'    Else
        '3.2 ��������� �������� ����, ����� ��������� � ������ ������ (����� 10 �����) � ������ ���� �� �������
        If Not extSquare Is Nothing Then
            '���� ������ ������������ � ������ ��������� ��������� �������
            If F_InsertFire.flag_StvolPlaceCheck Then
                Select Case extSquare.IsCellExtinguished(x, y)
                    Case Is = WaterExpenseKind.notSet               '���� �� �������� ��� ������ �� ��������� � ���� �������
                        '������ �� ������ - ���� ������ � ������ ��� �� � ��� �� ������
                    Case Is = WaterExpenseKind.notSufficient        '���� �������� ������������ - ������� ������ ������ ������ ���
                        If isOdd(step) And time > 10 Then
                            Exit Sub
                        End If
                    Case Is = WaterExpenseKind.sufficient           '���� ��������� ���������� - �� ������������ �����
                        Exit Sub
                End Select
            Else
                '���� ������ ������������ ��� ����� ��������� ��������� �������
                Select Case extSquare.IsSquareExtinguished(nozzleDemon, time)
                    Case Is = WaterExpenseKind.notSet               '���� �� �������� ��� ������ �� ��������� � ���� �������
                        '������ �� ������ - ���� ������ � ������ ��� �� � ��� �� ������
                    Case Is = WaterExpenseKind.notSufficient        '���� �������� ������������ - ������� ������ ������ ������ ���
                        If isOdd(step) And time > 10 Then
                            Exit Sub
                        End If
                    Case Is = WaterExpenseKind.sufficient           '���� ��������� ���������� - �� ������������ �����
                        Exit Sub
                End Select
            End If
        End If


    
    '4 ������� �������� ������
    '�� ���������
        RizeGettedCellPower x - 1, y - 1, firePower * diag
        RizeGettedCellPower x - 1, y + 1, firePower * diag
        RizeGettedCellPower x + 1, y - 1, firePower * diag
        RizeGettedCellPower x + 1, y + 1, firePower * diag
    '�� ����������
        RizeGettedCellPower x - 1, y, firePower * orto
        RizeGettedCellPower x + 1, y, firePower * orto
        RizeGettedCellPower x, y - 1, firePower * orto
        RizeGettedCellPower x, y + 1, firePower * orto
    
End Sub

Private Sub RizeGettedCellPower(ByVal x As Long, ByVal y As Long, ByVal value As Double)
'����������� ���������� �������� �������
    
    '���������, �� ������� �� ���������� ������ �� ������� �������
    If x < 0 Or y < 0 Or x > matrix.DimensionX Or y > matrix.DimensionY Then Exit Sub
    
    '���������, �� ������ �� ������ ������ � ������� �������
    If matrix.IsCellOutOfArea(x, y) Then Exit Sub
    
    '���������, �� �������� �� ��������� ������ ������
    If matrix.GetOpenSpaceLayerValue(x, y) = 1 Then Exit Sub
    
    '���� �� ����� ������ �� ���� ��������, ��������� �� � ��������� ������ ���������� ����� � ���� ����
    If matrix.GetPowerInOneStep(x, y) = 0 Then
        matrix.chargedCells.Add CreateNewCell(x, y)
    End If
    
    '����������� ��������
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, matrix.GetPowerInOneStep(x, y) + value
    
End Sub


Private Sub RizeCellFire(ByVal x As Long, ByVal y As Long)
'����������� �������� ������� ������
Dim startFirePower As Double
    startFirePower = matrix.GetActualFirePower(x, y)
    
    
    '���� ������ ��������� � �������� ���������� �� ���� ����� 0, �� �������
    If matrix.GetActualFirePower(x, y) = 0 And matrix.GetPowerInOneStep(x, y) = 0 Then Exit Sub

    '���� ������ ���������� ���� ���������� � ��� ���� ����������, �� ������ ����� ��������� �� � ��������� ������� ������
    If matrix.GetActualFirePower(x, y) = 0 And matrix.GetPowerInOneStep(x, y) > 0 Then
        matrix.firedCells.Add CreateNewCell(x, y)
        '��������� ������ � ��������� �������� ������
        AddToActiveCells x, y
    End If

    '����������� ������� �������� ������� ������ ������
    matrix.RizeCellFire x, y, matrix.GetPowerInOneStep(x, y)
    
    '���� ������ ����������, �� ��� ���� ��� ����� �� �� ������ ��������, ��������� �� � �������� ������������ ������
    If matrix.GetActualFirePower(x, y) > 0 And matrix.GetActualFirePower(x, y) < matrix.GetMaximumBurnPower Then
        matrix.startedFireCells.Add CreateNewCell(x, y)
    End If
    
    '��������� ������ � ��������� ������ ������ � ���������, ���� �������� ������� ������ �����������
    If startFirePower < matrix.GetMaximumBurnPower And matrix.GetFirePower(x, y) = matrix.GetMaximumBurnPower Then
        AddToFrontCells x, y
        AddToPerimeterCells x, y
        
        '��������� � ��������� ����������� ������� ������
        matrix.fireMaxCells.Add CreateNewCell(x, y), x & ":" & y
            
        '        '������������ �������
        '        Dim shp As Visio.Shape
        '        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
        '                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
        '        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
        '        shp.Cells("LinePattern").FormulaU = "0"
        '        shp.Cells("FillForegnd").FormulaU = "2"

'            DoEvents
    End If

    
    '������� �������� ���������� ������� �������
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, 0
End Sub

Private Sub RizeCellFireToMax(ByVal x As Long, ByVal y As Long)
'���������, ��� ������ ����� �� �������� (��������� ��� ��������� ��� �������� � ������������� �����)
    '���� ������ ���������� ���� ���������� � ��� ���� ����������, �� ������ ����� ��������� �� � ��������� ������� ������
    If matrix.GetActualFirePower(x, y) = 0 Then
        matrix.firedCells.Add CreateNewCell(x, y)
        '��������� ������ � ��������� �������� ������
'        AddToActiveCells x, y
    End If

    '����������� ������� �������� ������� ������ ������
    matrix.SetFireCell x, y
       
    '��������� ������ � ��������� ������ ������ � ���������, ���� �������� ������� ������ �����������
'    If startFirePower < matrix.GetMaximumBurnPower And matrix.GetFirePower(x, y) = matrix.GetMaximumBurnPower Then
        
        AddToFrontCells x, y
        AddToPerimeterCells x, y
        AddToActiveCells x, y
            
'        '������������ �������
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'        shp.Cells("LinePattern").FormulaU = "1"
'        shp.Cells("LineColor").FormulaU = "2"
'        shp.Cells("FillPattern").FormulaU = "0"

'            DoEvents
'    End If

    
    '������� �������� ���������� ������� �������
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, 0
End Sub

Private Sub AddToActiveCells(ByVal x As Long, ByVal y As Long)
'��������� ������ � ��������� �������� ������. ���� �� ������ � ��������� ������ ��� ������� - �������
    On Error Resume Next
    matrix.activeFiredCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromActiveCells(ByVal x As Long, ByVal y As Long)
'������� ������ �� ��������� �������� ������. ���� �� ������ � ��������� ������ ����������� - �������
    On Error Resume Next
    matrix.activeFiredCells.Remove x & ":" & y
End Sub

Private Sub AddToFrontCells(ByVal x As Long, ByVal y As Long)
'��������� ������ � ��������� ������ ������. ���� �� ������ � ��������� ������ ��� ������� - �������
    On Error Resume Next
    matrix.fireFrontCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromFrontCells(ByVal x As Long, ByVal y As Long)
'������� ������ �� ��������� ������ ������. ���� �� ������ � ��������� ������ ����������� - �������
    On Error Resume Next
    matrix.fireFrontCells.Remove x & ":" & y
End Sub

Private Sub AddToPerimeterCells(ByVal x As Long, ByVal y As Long)
'��������� ������ � ��������� ������ ���������. ���� �� ������ � ��������� ������ ��� ������� - �������
    On Error Resume Next
    matrix.firePerimeterCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromPerimeterCells(ByVal x As Long, ByVal y As Long)
'������� ������ �� ��������� ������ ���������. ���� �� ������ � ��������� ������ ����������� - �������
    On Error Resume Next
    matrix.firePerimeterCells.Remove x & ":" & y
End Sub

Public Sub RefreshFirePerimeter()
'��������� ��� ������ ����������� �� ������� ���� ������� � �������� ���������
'��� ��������� ����������� ������ ��� ��������� ��� ��� ��� ��������� ������������ ��������� ������������
Dim cell As c_Cell
    
    '��������� ������, �� ������� ����, �� ����� �� ��� �����������, � � ���� ������ ������� �� �� ���������
    For Each cell In matrix.firedCells
        If Not matrix.IsInner(cell.x, cell.y) Then
            AddToActiveCells cell.x, cell.y
            AddToFrontCells cell.x, cell.y
        End If
    Next cell
End Sub


'-----------------------------��������---------------------------------------------
Public Property Get CurrentStep()
'������� ��� �������
    CurrentStep = step
End Property
Public Property Get CalculatedStep()
'������������ ��� �������
    CalculatedStep = calcStep
End Property

Public Sub RizeCurrentStep()
'����������� ������� ��� �������
    step = step + 1
End Sub



'-----------------------------�����������-------------------------------------------
Private Function CreateNewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'������� ����� ������
    On Error GoTo EX

Dim cell As c_Cell
    
    Set cell = New c_Cell
    Set cell.ParentMatrix = matrix
    cell.x = x
    cell.y = y
    
Set CreateNewCell = cell
Exit Function
EX:
     Debug.Print "err"
End Function



'-----------------------------������-------------------------------------------
Public Function GetFiredCellsCount() As Long
    GetFiredCellsCount = matrix.firedCells.Count
End Function
Public Function GetActiveCellsCount() As Long
    GetActiveCellsCount = matrix.activeFiredCells.Count
End Function

Public Function GetFireSquare() As Single
'���������� ������� ������
    GetFireSquare = matrix.fireMaxCells.Count * grain ^ 2 / 1000000
End Function

Public Function GetExtSquare() As Single
'���������� ������� �������
Dim tmpVal As Single
    If Not extSquare Is Nothing Then
        tmpVal = extSquare.GetSquare
    Else
        tmpVal = 0
    End If
    
    If tmpVal = 0 Then
        tmpVal = GetFireSquare
    End If
    GetExtSquare = tmpVal
End Function

Public Function GetOpenSpaceLayer() As Variant
    GetOpenSpaceLayer = matrix.GetOpenSpaceLayer
End Function

'----------------------------������ �������� �������--------------------------
Public Sub FireSquareCalculate2()
'��� ������� ������� ������� ��� ����� "�� ��������� �������� ��������� �������"
Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
    
    Set extSquare = New c_ExtinguishingSquare
    extSquare.intenseNeed = intenseNeed
    
    Set frontDemonImpl = New c_FrontDemon
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    frontDemonImpl.CreateMatrix matrix.DimensionX, matrix.DimensionY
    frontDemonImpl.grain = grain
    frontDemonImpl.SetStartState matrix.fireFrontCells
    frontDemonImpl.RunDemon2 extSquare, matrix.fireFrontCells, matrix.fireFrontCells
    
    extinguishingSquareDemon.CreateMatrix matrix
    extinguishingSquareDemon.setGrain = grain
    
    '������������ ������ ������� �� ������
    extinguishingSquareDemon.RunDemon extSquare
    
    '����� ����������
    Dim blck As c_ExtinguishingBlock
    Dim extSquareMeters As Single
    Dim needExpense As Single
    Dim factExpense As Single
    Dim expKindForAll As String
    For Each blck In extSquare.extinguishingBlocks
        extSquareMeters = extSquareMeters + blck.GetSquare
        needExpense = needExpense + blck.GetExpenceNeed
        factExpense = factExpense + blck.expenseFact
        expKindForAll = expKindForAll & " " & blck.GetWaterExpenseKind
    Next blck
    

        Debug.Print "������� �������: " & extSquareMeters
        Debug.Print "��������� ������ : " & needExpense
        Debug.Print "����������� ������ : " & factExpense
        Debug.Print "������������� �������: " & expKindForAll

End Sub

Public Sub FireSquareCalculate(ByRef nozzleCollection As Collection)
'��������� ��������� ������ ������� �������
Dim nozzle As c_FireNozzleModel
Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
    
    Set extSquare = New c_ExtinguishingSquare
    extSquare.intenseNeed = intenseNeed
    
    Set frontDemonImpl = New c_FrontDemon
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    frontDemonImpl.CreateMatrix matrix.DimensionX, matrix.DimensionY
    frontDemonImpl.grain = grain
    frontDemonImpl.SetStartState matrix.fireFrontCells
    
    extinguishingSquareDemon.CreateMatrix matrix
    extinguishingSquareDemon.setGrain = grain
    
    
    Dim i As Integer
    i = 0
    For Each nozzle In nozzleCollection
        If Not nozzle.fireContactCell Is Nothing Then
'            Debug.Assert i < 1
            
'            If i > 0 Then frontDemonImpl.DrawBlockExists extSquare, nozzle.fireContactCell
            '��������� ��������� ���������� ������� ������ ������ (� ������ ������� �� �� ������)
            frontDemonImpl.RunDemon extSquare, matrix.fireFrontCells, nozzle
            
'            '������������ ������ ������� �� ������
'            extinguishingSquareDemon.RunDemon extSquare

            i = i + 1
        End If
    Next nozzle
    
    '������������ ������ ������� �� ������
    extinguishingSquareDemon.RunDemon extSquare
    
    Dim blck As c_ExtinguishingBlock
    Dim extSquareMeters As Single
    Dim needExpense As Single
    Dim factExpense As Single
    Dim expKindForAll As String
    For Each blck In extSquare.extinguishingBlocks
        extSquareMeters = extSquareMeters + blck.GetSquare
        needExpense = needExpense + blck.GetExpenceNeed
        factExpense = factExpense + blck.expenseFact
        expKindForAll = expKindForAll & " " & blck.GetWaterExpenseKind
    Next blck
    
'    If extSquare.extinguishingBlocks.Count > 0 Then
'        Set blck = extSquare.extinguishingBlocks(1)
        Debug.Print "������� �������: " & extSquareMeters
        Debug.Print "��������� ������ : " & needExpense
        Debug.Print "����������� ������ : " & factExpense
        Debug.Print "������������� �������: " & expKindForAll
'    End If
    
   
End Sub

Public Sub DrawExtSquareByDemon(ByRef fireShp As Visio.Shape)
'������� ������ ���������� ������� �������
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
Dim extShp As Visio.Shape
Dim ExtSquareFormula As String
Dim squareShapesCol As Collection
   
    '��������� ������� �� ����������� �������� �������
    If extSquare.extinguishingBlocks.Count = 0 Then
        Exit Sub
    End If
    
    ClearLayer "ExtSquare"
    
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon

    extinguishingSquareDemon.setGrain = grain
    Set squareShapesCol = extinguishingSquareDemon.DrawFireSquare(extSquare)
    
    '���������� ����� ������� ��� ������ �� ������� ���������� �������� �������
    For Each extShp In squareShapesCol
        If ExtSquareFormula = "" Then
            ExtSquareFormula = ExtSquareFormula & "IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        Else
            ExtSquareFormula = ExtSquareFormula & "+IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        End If
    Next extShp
    
    '����������� �������� ������� ������� ��������������� ������ ������� ������ �������� ���� �������� ������� ��������� ����� (�������)
    fireShp.Cells("Prop.ExtFull").FormulaU = "Index(1, Prop.ExtFull.Format)"
    fireShp.Cells("Prop.ExtSquareT").FormulaU = ExtSquareFormula
End Sub

Public Function GetWaterExpenseKind() As WaterExpenseKind
'��������� ������������� ������� ���� ��� �������
Dim tmpVal As WaterExpenseKind
    
    If Not extSquare Is Nothing Then
        tmpVal = extSquare.GetWaterExpenseKind
    Else
        tmpVal = WaterExpenseKind.notSufficient
    End If
    
GetWaterExpenseKind = tmpVal
End Function

Public Sub ClearFireExtSquare()
    Set extSquare = Nothing
End Sub

Public Sub NozllesRecalculate()
    '��������� ������ ������� ��������� �������
    nozzleDemon.RunDemon time

    '������������ ������� ������� ��� ������� ���������� �� �������� ������, ��� �������, ��� ��� ���� ������ ����� �������� � ������� ������
    If F_InsertFire.flag_StvolPlaceCheck Then
        FireSquareCalculate nozzleDemon.nozzles
    Else
        FireSquareCalculate2
    End If
End Sub

''-------�������� ���������
'Public Sub DrawActiveCells()
''������������ �������
'Dim cell As c_Cell
'
'    For Each cell In matrix.activeFiredCells
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("�����")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'    Next cell
'End Sub
'
'Public Sub DrawFrontCells()
''������������ �������
'Dim cell As c_Cell
'
'    For Each cell In matrix.fireFrontCells
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("�����")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'    Next cell
'End Sub
'
'Public Sub RemoveActive()
'Dim cell As c_Cell
'
'    For Each cell In matrix.activeFiredCells
'        If matrix.IsInner(cell.x, cell.y) Then
'            RemoveFromActiveCells cell.x, cell.y
'            RemoveFromFrontCells cell.x, cell.y
'        End If
'    Next cell
'End Sub

Public Sub DrawPerimeterCells()
'������������ ������ ��������� ��� ������ ������
Dim perimterDemon As c_PerimeterDemon
Dim fireConturs As Visio.Selection
Dim shp As Visio.Shape
    
    ClearLayer "Fire"
    
    Set perimterDemon = New c_PerimeterDemon
    perimterDemon.setGrain(grain).RunDemon matrix.firePerimeterCells
    perimterDemon.DrawSquare
    
    '���������� ���������� ������ � ������������� �� ������� ����
'    On Error Resume Next
    Set fireConturs = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    If fireConturs.Count > 0 Then
        fireConturs.Combine
        Set shp = Application.ActiveWindow.Selection(1)
        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
    End If
    
End Sub
