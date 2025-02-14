VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireNozzlesDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public nozzles As Collection                        '��������� �������� ������� ����������� �� �����

Dim matrix As Variant                               '������� ��������� ����
Dim calculatedCells As Variant                      '������� ������������ ������
Private fireMatrixObject As c_Matrix                 '������ ��������� ���� (�������, ������)

Private grain As Integer                             '����� �������

Private form As Object                              '����� ������ ������

Const halfRadian As Double = 1.570796326
Const diag As Double = 1.4142135623731              '����������� ��������������� ����� � ����
Const orto As Double = 1                            '����������� ��������������� �� ���������

Const mmInInch As Double = 25.4

'------------------����� ������ ���������� �������----------------------------
'1 - ������� ��� ������ ������������� �� �����
'2 - ������ ���� ������������ ��� ������� (����������)
'3 - ������ �� ����� ������������ ������� ������ ������� �� ������ �������������� ������ ������� �������



Private Sub Class_Initialize()
    Set nozzles = New Collection
End Sub

Public Property Let setGrain(ByVal grainValue)
    grain = grainValue
End Property

Public Property Set SetMatrix(ByRef matr As c_Matrix)
    Set fireMatrixObject = matr
End Property

Public Function SetForm(ByRef frm As Object) As c_FireNozzlesDemon
    Set form = frm
Set SetForm = Me
End Function




Public Sub RunDemon(ByVal currentTime As Single)
'��������� ������
Dim nozzle As c_FireNozzleModel
       
    '���������� ��� ������ � ���������, �������� �� �� ���� ������� (5� ��� ������ � 10 ��� ��������) ������ ������
    For Each nozzle In nozzles
        ReDim matrix(fireMatrixObject.DimensionX, fireMatrixObject.DimensionY)
        ReDim calculatedCells(fireMatrixObject.DimensionX, fireMatrixObject.DimensionY)
        
        '���������, ��������� �� ����� � �������� ��������
        If nozzle.position.x < 0 Or nozzle.position.y < 0 Or _
                nozzle.position.x > UBound(calculatedCells, 1) Or nozzle.position.y > UBound(calculatedCells, 2) Then
        Else
            '��������� ���� "��������� ������ ������ ����� �� ������"
            If F_InsertFire.flag_UseStvolsAfterTime Then
                '��������� ��� �� ����� �����
                If currentTime >= nozzle.timeOfSet Then
                    calculatedCells(nozzle.position.x, nozzle.position.y) = 1
                    '���� ���� ������������ ������ �������� ���-���� ������, ����� ������ ������� ��������� �� ������ ����� �������
                    checkFireIntercourse nozzle
                End If
            Else
                calculatedCells(nozzle.position.x, nozzle.position.y) = 1
                '���� ���� ������������ ������ �������� ���-���� ������, ����� ������ ������� ��������� �� ������ ����� �������
                checkFireIntercourse nozzle
            End If
            

        End If
        

    Next nozzle
    
    
    

End Sub




Public Sub FindNozzles()
'��������� ��������� �������
Dim shp As Visio.Shape
Dim nozzle As c_FireNozzleModel
Dim indexPers As Integer
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            indexPers = shp.Cells("User.IndexPers").Result(visNumber)
            If indexPers = 34 Or indexPers = 36 Or indexPers = 39 Then      '34 - ������ ������� �����, 36 - �������� �������, 39 - ������� �������
                '�������� �� �������������
                If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
                    '�������� �� ������ �� ������� ��� ������
                    If (form.flag_StvolOnlyForExtinguishing And ShapeHaveCell(shp, "Prop.UseDirection", "�������;������")) _
                        Or Not form.flag_StvolOnlyForExtinguishing Then
                            Set nozzle = New c_FireNozzleModel
                            If indexPers = 34 Then      '34 - ������ ������� �����, 36 - �������� �������, 39 - ������� �������
                                nozzle.nozzleType = waterHand
                            ElseIf indexPers = 36 Or indexPers = 39 Then
                                nozzle.nozzleType = waterLafet
                            End If
                            Set nozzle.position = GetNozzlePosition(shp.Cells("PinX").Result(visMillimeters), _
                                shp.Cells("PinY").Result(visMillimeters))
                            nozzle.waterExpense = shp.Cells("User.PodOut").Result(visNumber)
                            nozzle.angle = shp.Cells("Angle").Result(visNumber) + halfRadian
                            nozzle.timeOfSet = getNozzleTime(shp)
                            
                            
                            nozzles.Add nozzle
                    End If
                    
                    

                End If
            End If
        End If
    Next shp
End Sub

Public Function GetExpense(ByVal currentTime As Single) As Single
'���������� ������ ���� �������
Dim nozzle As c_FireNozzleModel
    
    For Each nozzle In nozzles
        If currentTime >= nozzle.timeOfSet Then    '�� ��������� ����� ������ ������
            GetExpense = GetExpense + nozzle.waterExpense
        End If
    Next nozzle
    
End Function


Private Function GetNozzlePosition(ByVal xPos As Double, ByVal yPos As Double) As c_Cell
'������� ���������� ������������ ������
Dim x As Integer
Dim y As Integer

    x = Int(xPos / grain)
    y = Int(yPos / grain)
    
    Set GetNozzlePosition = NewCell(x, y)
End Function

Private Function getNozzleTime(shp As Visio.Shape) As Single
Dim t1 As Date
Dim t2 As Date

    t1 = Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").ResultStr(visUnitsString)
    t2 = shp.Cells("Prop.SetTime").ResultStr(visUnitsString)
    getNozzleTime = DateDiff("n", t1, t2)
    
End Function


Private Sub checkFireIntercourse(ByRef nozzle As c_FireNozzleModel)
'��������� ����������� ����������� ������ ������ � �������
Dim coll As Collection
Dim maxWay As Integer

Dim cs As Double
Dim sn As Double
Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim cell As c_Cell

    If F_InsertFire.optValue Then
        On Error GoTo EX
'        maxWay = F_InsertFire.txtNozzleRangeValue * 1000
        maxWay = form.StvolCalcDistance
        On Error GoTo EX1
    Else
        If nozzle.nozzleType = waterLafet Then
            maxWay = 10000
        Else
            maxWay = 5000
        End If
    End If
    
'    Set coll = New Collection
'    coll.Add nozzle.position
'
'    Do While coll.Count > 0
'        Set coll = OneStep(nozzle, coll, maxWay)
'        If coll Is Nothing Then Exit Do
'    Loop
    

    '��������� ����������� �� ����������� ����������� �� �������
        '������� ���������� ������� � ����� ���� - ��� ����������� ���������� �����
        cs = Cos(nozzle.angle)
        sn = Sin(nozzle.angle)
        
        '��������������� ��������� ��� ������� �����, �� ������� ����, ������������ �� ��� � ������� ������
        i = 0
        Set cell = New c_Cell
        cell.x = nozzle.position.x
        cell.y = nozzle.position.y
        Do While CellType = 0
            cell.x = nozzle.position.x + Int(cs * i)
            cell.y = nozzle.position.y + Int(sn * i)
'------------------
'            Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("spray")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'------------------
            
'            Debug.Print cell.x & ", " & cell.y & " -> " & nozzle.position.x & ", " & nozzle.position.y
            
            If IsIntercourseWithFront2(nozzle, cell) Then
                Exit Do
            End If
            If Math.Sqr((cell.x - nozzle.position.x) ^ 2 + (cell.y - nozzle.position.y) ^ 2) * grain > maxWay Then
                Exit Do
            End If
            i = i + 1
        Loop
    
Exit Sub
EX:
    MsgBox "�� ��������� ������� �������� ���������� �������� �������"
Exit Sub
EX1:
    
End Sub


Private Function OneStep(ByRef nozzle As c_FireNozzleModel, ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'������� ���������� ������, ���� ������ �������� ������� ������ ������ (����� ����� ���� ���� ������� �� ��������� ������)
'�� ����� �������� ��������� ������� ����� ������, �� ������ ��������� ����� ������ ���� ��������������� �����
Dim cell As c_Cell
Dim newColl As Collection
Dim cellWay As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
'        cellWay = matrix(cell.x, cell.y)
'        '��������������� �� ������
'            '������� ���������� ������� � ����� ���� - ��� ����������� ���������� �����
'            cs = Cos(nozzle.angle)
'            sn = Sin(nozzle.angle)
'
'            '��������������� ��������� ��� ������� �����, �� ������� ����, ������������ �� ��� � ������� ������
'            x = cell.x + Int(cs * i)
'            y = cell.y + Int(sn * i)
'            Set cell = NewCell(x, y)
'
'            If IsIntercourseWithFront(nozzle, newColl, cell, _
'                                        Math.Sqr((x - nozzle.position.x) ^ 2 + (y - nozzle.position.y) ^ 2) * grain, maxWay) Then
'                Set OneStep = Nothing
'                Exit Function
'            End If
        
        '�������������� �� �������� ������
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, l), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, lu), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, u), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, ru), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, r), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, rd), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, d), cellWay + orto * grain, maxWay) Then
        Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, ld), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If

    Next cell
    
    Set OneStep = newColl
    
    

End Function

Private Function IsIntercourseWithFront(ByRef nozzle As c_FireNozzleModel, ByRef coll As Collection, ByRef cell As c_Cell, _
                                        ByVal wayLen As Single, ByVal maxWay As Integer) As Boolean
'��������� ���� ���������� �� ������ � ���� �� ������ ���������, �������� ���
'������ �������� �������� � ��������� ���������
    '���� ������ ��������� �� ��������� ������� - �������
    If Not IsInMatrix(cell) Then Exit Function
    
    '���������, �������� �� ������ �������� �������������
    If Not fireMatrixObject.GetOpenSpaceLayerValue(cell.x, cell.y) = CellState.csOpenSpace Then Exit Function
    
    '���� ���� ��������� ���������� �������� - �� ��������� ���
    If wayLen > maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Function
    
    '���� ������ ������������� ����, ��...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        '���������, �������� �� ������ ������� ������ ������
    '�������� �������� ������������� �������!!!
        If fireMatrixObject.IsCellMaxFire(cell.x, cell.y) Then
            If Not fireMatrixObject.IsInner(cell.x, cell.y) Then
                '��������� ��� ������ �������� ����� �������� � ������� ������
'                Debug.Print "cell." & cell.x & ":" & cell.y & " ������� �� ���������!!!"
                Set nozzle.fireContactCell = cell
                
                '������� �� ������� ��������� ������ - ���� ������ ������ �������� ������ ������ (�������!)
                IsIntercourseWithFront = True
                Exit Function
            End If
        End If
        
        '������ ����� �������� � �������� �������� � ��������� (���� ����� ������ � ��������� ��� ����, ������ �� ���������� - ����� Error)
        matrix(cell.x, cell.y) = wayLen

        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
    
End Function

Private Function IsIntercourseWithFront2(ByRef nozzle As c_FireNozzleModel, ByRef cell As c_Cell) As Boolean
'��������� �������� �� ������ ������ ������ ������
    '���� ������ ��������� �� ��������� ������� - �������
    If Not IsInMatrix(cell) Then Exit Function
    
    '���������, �������� �� ������ �������� �������������
    If Not fireMatrixObject.GetOpenSpaceLayerValue(cell.x, cell.y) = CellState.csOpenSpace Then
        '������� �� ������� - ������� �����, �� ����� �� ������, ������ ����� ������� ������ �� ������
        IsIntercourseWithFront2 = True
        Exit Function
    End If
    
'    '���� ������ ������������� ����, ��...
'    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        '���������, �������� �� ������ ������� ������ ������
    '�������� �������� ������������� �������!!!
        If fireMatrixObject.IsCellMaxFire(cell.x, cell.y) Then
            If Not fireMatrixObject.IsInner(cell.x, cell.y) Then
                '��������� ��� ������ �������� ����� �������� � ������� ������
'                Debug.Print "cell." & cell.x & ":" & cell.y & " ������� �� ���������!!!"
                Set nozzle.fireContactCell = cell
                
                '������� �� ������� ��������� ������ - ���� ������ ������ �������� ������ ������ (�������!)
                IsIntercourseWithFront2 = True
                Exit Function
            End If
        End If
        
'        '������ ����� �������� � �������� �������� � ��������� (���� ����� ������ � ��������� ��� ����, ������ �� ���������� - ����� Error)
'        matrix(cell.x, cell.y) = wayLen

'        On Error Resume Next
'        coll.Add cell, cell.x & ":" & cell.y
'    End If
    
End Function




'-----------------------�����������-----------------------------------------
Private Function NextCell(ByRef cell As c_Cell, ByVal Dir As Directions) As c_Cell
'���������� ��������� ������ �������� ���������� �����������
    Select Case Dir
        Case Is = Directions.l
            Set NextCell = NewCell(cell.x - 1, cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(cell.x - 1, cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(cell.x, cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(cell.x + 1, cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(cell.x + 1, cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(cell.x + 1, cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(cell.x, cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(cell.x - 1, cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'������� ����� ������
Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    Set NewCell = cell
End Function

Private Function IsInMatrix(ByRef cell As c_Cell) As Boolean
'���������� ������, ���� ������ ��������� � �������� �������
    If cell.x <= 0 Or cell.y <= 0 Or cell.x > UBound(matrix, 1) Or cell.y > UBound(matrix, 2) Then
        IsInMatrix = False
    Else
        IsInMatrix = True
    End If
End Function


