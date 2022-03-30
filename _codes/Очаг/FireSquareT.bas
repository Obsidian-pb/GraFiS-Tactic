Attribute VB_Name = "FireSquareT"
Public fireModeller As c_Modeller
Dim frmF_InsertFire As F_InsertFire
Public grain As Integer

Public stopModellingFlag As Boolean      '���� ��������� �������������

'------------------------������ ��� ���������� ������� ������ � �������������� ������������ ������-------------------------------------------------

Public Sub MakeMatrix(ByRef controlForm As Object)
'��������� �������
Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder
    

    '---���������� ������
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    

    
    '�������� ������� �������� �����������
    Set matrixBuilder = New c_MatrixBuilder
    matrixBuilder.SetForm controlForm
    matrix = matrixBuilder.NewMatrix(grain)

    '���������� ������ �������
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

    '���������� ���������
    Set fireModeller = New c_Modeller
    fireModeller.SetMatrix matrixObj
    
    '��������� ��������� �������� �����
    fireModeller.grain = grain

    '���� ������ ����� � �� �� ����������� ������������� ����� ������ ������
    GetFirePoints
    
    controlForm.lblMatrixIsBaked.Caption = "������� �������� �� " & tmr.GetElapsedTime & " ���."
    controlForm.lblMatrixIsBaked.ForeColor = vbGreen
    
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub

Public Sub RefreshOpenSpacesMatrix(ByRef controlForm As Object)
'��������� ������� �������� �����������
Dim matrix() As Variant
Dim matrixBuilder As c_MatrixBuilder
    
    If fireModeller Is Nothing Then
        MsgBox "�� �� ������ �������� �� ���������� �������!", vbCritical
        Exit Sub
    End If
    
    '---���������� ������
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    
    '�������� ������� �������� �����������
    Set matrixBuilder = New c_MatrixBuilder
    matrixBuilder.SetForm controlForm
    matrix = matrixBuilder.NewMatrix(grain)
    
    '��������� ������� �������� �����������
    fireModeller.refreshOpenSpaces matrix
    
    '��������� �������� ������
    fireModeller.RefreshFirePerimeter
    
    '������� ��������� � ������ ����������
    controlForm.lblMatrixIsBaked.Caption = "������� ��������� �� " & tmr.GetElapsedTime & " ���."
    controlForm.lblMatrixIsBaked.ForeColor = vbGreen

    tmr.PrintElapsedTime
    Set tmr = Nothing
    
End Sub


Public Sub RunFire(ByVal timeElapsed As Single, ByVal speed As Single, ByVal intenseNeed As Single, Optional ByVal path As Single)
'���������� ������� ������� �� ��� ���, ���� ��������� ���� ���������� ����� �� ������ ������ distance + ���������� ����� (�������� � ���������)
Dim vsO_FireShape As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim newFireShape As Visio.Shape
Dim modelledFireShape As Visio.Shape
Dim borderShape As Visio.Shape

    '�������� ���������� ������ - ��� �������������� �� ���������� ���������� �������
    On Error GoTo ex
    
    '���� ���� ����� 0, �� ��������� ��� ���������� �������
    If path = 0 Then path = 10000
    
    '---���������� ������
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    i = 1
    
    '---���������� ���������� �������� ����������� ���� (���� ������� ����� + ���� ���������� �����)
    Dim boundDistance As Single             '���������� ����������, �������� �������
    Dim currentDistance As Single           '������� ���������� ����������
    Dim prevDistance As Single              '���������� ���������� �� ���������� ����� �������
    Dim diffDistance As Single              '���������� ���������� � ������ ����� �������
    Dim realCurrentDistance As Single       '�������� ������� ���������� ����������
    Dim realDiffDistance As Single          '�������� ���������� ���������� � ������ ����� �������
    Dim currentTime As Single               '������� ����� � ������ �������
    Dim prevTime As Single                  '����� �� ������� �������� ���������� ���� �������
    Dim diffTime As Single                  '����� �� ������� �������� ������� ���� �������
    
    '---Activate nozzles calculations
    fireModeller.ActivateNozzles F_InsertFire
    
    '��������� ��������� �������� ��������� ������������� ������ ����
    fireModeller.intenseNeed = intenseNeed
    
    
    prevDistance = fireModeller.distance
    boundDistance = timeElapsed * speed + prevDistance
    
    prevTime = fireModeller.time
    
    Do While diffTime < timeElapsed And realCurrentDistance < path
        ClearLayer "ExtSquare"
        
'        Stop   ' - ����� ����� �������� �������� �� ������������� ������� ��� ������� -> fireModeller.GetExtSquare
        If fireModeller.GetExtSquare < fireModeller.GetFireSquare Then
'        fireModeller.
            '���������, ������� ������� ������ ������, ���� ������ 10 �����, �� �����������, ������ ������ ������ ���, �.�., � ��������� ��������
            If currentTime < 10 Then
                '��� ������� ����� 10 ����� ������� ���� ������ ������ ������ ���
                If IsEven(fireModeller.CurrentStep) Then
                    fireModeller.OneRound
    
                    '���������� ����������� ����� � ���� ������
                    MakeShape
                End If
            Else
                fireModeller.OneRound
                    
                '���������� ����������� ����� � ���� ������
                MakeShape
            End If
        ElseIf fireModeller.GetExtSquare >= fireModeller.GetFireSquare Then
            If Not fireModeller.GetWaterExpenseKind = sufficient Then   '���� ���������� ������� �� ������ �� ������, ������ ������� ��������� ���
            
'            Else
                '���������, ������� ������� ������ ������, ���� ������ 10 �����, �� �����������, ������ ������ ������ ���, �.�., � ��������� ��������
                If currentTime < 10 Then
                    '��� ������� ����� 10 ����� ������� ���� ������ ������ ������ ���
                    If IsEven(fireModeller.CurrentStep) Then
                        fireModeller.OneRound
        
                        '���������� ����������� ����� � ���� ������
                        MakeShape
                    End If
                Else
                    fireModeller.OneRound
                        
                    '���������� ����������� ����� � ���� ������
                    MakeShape
                End If
            End If
        End If
        
        '����������� ��� �������
        fireModeller.RizeCurrentStep
            
        currentDistance = GetWayLen(fireModeller.CurrentStep, grain)
        diffDistance = currentDistance - prevDistance
        realCurrentDistance = GetWayLen(fireModeller.CalculatedStep, grain)
        realDiffDistance = realCurrentDistance - prevDistance
        
        currentTime = currentDistance / speed
        diffTime = currentTime - prevTime
               
        On Error Resume Next
        '---�������� ������� ������������� �������
        F_InsertFire.lblCurrentStatus.Caption = "���: " & i & "(" & fireModeller.CurrentStep & "), " & _
                                                " ���������� ����: " & Round(realDiffDistance, 2) & "(" & Round(realCurrentDistance, 2) & ")�.," & _
                                                " �����: " & Round(diffTime, 2) & "(" & Round(currentTime, 2) & ")���, " & _
                                                Chr(13) & "������� ������: " & fireModeller.GetFireSquare & "�.��., " & _
                                                Chr(13) & "������� �������: " & fireModeller.GetExtSquare & "�.��., " & _
                                                Chr(13) & "��������� ������: " & fireModeller.GetExtSquare * fireModeller.intenseNeed & "�/�"
        '��������� ����� �������� ����� ��������� � ������ �������������
        F_InsertFire.timeElapsedMain = currentTime
        '��������� ����� �������� ���� ���������� � ������ �������������
        F_InsertFire.pathMain = realCurrentDistance
        
        
        On Error GoTo ex
        
        i = i + 1
        
        fireModeller.distance = realCurrentDistance ' currentDistance
        fireModeller.time = currentTime
               
        '������� ��������� � ��������� ������� ������������
        Application.ActiveWindow.DeselectAll
        DoEvents
        
        '���� ������������ ����� � ����� ������ "����������" ���������� �������������
        If stopModellingFlag Then
            Exit Do
        End If
    Loop
        
    '---���������� ������������ ������ � �������� �� � ������ ������� �������
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    Set modelledFireShape = vsoSelection(1)
    Application.ActiveWindow.Select modelledFireShape, visSelect
    
    '---���������� ���������
    ImportAreaInformation
'    '---��������� ��� ������ ����������� ������� �������
    If fireModeller.GetExtSquare > 0 And F_InsertFire.flag_DrawExtSquare.value = True Then
        fireModeller.DrawExtSquareByDemon modelledFireShape
    End If
    '���������� ���������� ������ �� ������ ����
    modelledFireShape.SendToBack
    
    '���������� ������ ��������� ���� (��� �� �������) �� ������ ����
    If TryGetShape(borderShape, "User.IndexPers:1001") Then
        borderShape.SendToBack
    End If
        
''TEST:
'fireModeller.DrawExtSquareByDemon modelledFireShape
'������ ����� �� ����������� ����� ������ ���� �������
Application.ActiveWindow.DeselectAll
Application.ActiveWindow.Select modelledFireShape, visSelect

        
    Debug.Print "����� ��������� " & tmr2.GetElapsedTime & "�."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
    
Exit Sub
ex:
    MsgBox "������� �� ��������!", vbCritical
    
    '---���������� ������������ ������ � �������� �� � ������ ������� �������
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    Set modelledFireShape = vsoSelection(1)
    Set newFireShape = ActivePage.Drop(modelledFireShape, _
                        modelledFireShape.Cells("PinX").Result(visInches), modelledFireShape.Cells("PinY").Result(visInches))
    
    '---���������� ���������
    ImportAreaInformation
    '���������� ���������� ������ �� ������ ����
    newFireShape.SendToBack
    modelledFireShape.SendToBack
        
    Debug.Print "����� ��������� " & tmr2.GetElapsedTime & "�."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
End Sub

' ����������� ������� (������� ������)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
End Sub

Public Function IsAcceptableMatrixSize(ByVal maxMatrixSize As Long, ByVal grain As Integer) As Boolean
Dim xCount As Long
Dim yCount As Long
Dim shp As Visio.Shape

    On Error GoTo ex
    
    '��������� ��� �� �� ������ �������� ������ ��������� ����. ���� ����, �� ����������, ��� ������ ��������
    If TryGetShape(shp, "User.IndexPers:1001") Then
        IsAcceptableMatrixSize = True
        Exit Function
    End If
'    grain = Me.txtGrainSize.value

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
    
    IsAcceptableMatrixSize = xCount * yCount < maxMatrixSize
Exit Function
ex:
    IsAcceptableMatrixSize = False
End Function




'�� ����� ������ ���
'Public Sub DrawExtSquare()
''������� ������� �� ��������� ������� �������
'    fireModeller.DrawExtSquareByDemon
'End Sub









Private Sub GetFirePoints()
'������ ���� � ��������� ����� ������ �������
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers") = 70 Then
                '---������������� ���������� �����, ��� ����������� ������� ��������������� ����
                SetFirePointFromCoordinates shp.Cells("PinX").Result(visMillimeters), _
                    shp.Cells("PinY").Result(visMillimeters)
            End If
        End If
    Next shp
   
End Sub

Private Sub SetFirePointFromCoordinates(xPos As Double, yPos As Double)
'�������� � ������� ������� ������ �� ��������� �������������� �����������
Dim xIndex As Integer
Dim yIndex As Integer

    xIndex = Int(xPos / grain)
    yIndex = Int(yPos / grain)
    
    fireModeller.SetStartFireCell xIndex, yIndex

End Sub

Private Sub MakeShape()
'������������ ������ ���� ������� ��� ������ ������
    fireModeller.DrawPerimeterCells
End Sub

Public Function GetStepsCount(ByVal grain As Integer, ByVal speed As Single, ByVal elapsedTime As Single) As Integer
'������� ���������� ���������� ����� � ����������� �� ������� �����, �������� ��������������� ���� � ������� �� ������� ������������ ������

    '1 ���������� ���� ������� ������ ������ �����
    Dim firePathLen As Double
    firePathLen = speed * elapsedTime * 1000 / grain
    
    '2 ���������� ���������� ������� ����� ����� ��� ����������
    Dim tmpVal As Integer
    tmpVal = firePathLen / 0.58

    GetStepsCount = IIf(tmpVal < 0, 0, tmpVal)
    
End Function

Public Function GetWayLen(ByVal stepsCount As Integer, ByVal grain As Double) As Single
'������� ���������� ���������� ���� � ������
    Dim metersInGrain As Double
    metersInGrain = grain / 1000

    GetWayLen = CalculateWayLen(stepsCount) * metersInGrain
End Function

Public Function CalculateWayLen(ByVal stepsCount As Integer) As Integer
'������� ���������� ���������� ���� � �������
    Dim tmpVal As Integer
    tmpVal = 0.58 * stepsCount
    CalculateWayLen = IIf(tmpVal < 0, 0, tmpVal)
End Function




Public Function IsMatrixBacked() As Boolean
'���������� ������, ���� ������� ��� �������� � ����, ���� ���
    IsMatrixBacked = Not fireModeller Is Nothing
End Function

Private Function IsEven(ByVal number As Integer) As Boolean
'���������, ������ �� �����
    IsEven = Int(number / 2) = number / 2
End Function


'------------------------------------���������� � ������� ������ ��������� ������------------------------------
Public Sub AddFireArea(ShpObj As Visio.Shape)
'���������� � ������� ������ ��������� ������
        
    If Not IsMatrixBacked Then
        MsgBox "������� �� ��������!!!"
        Exit Sub
    End If
    
    '����������� ������ � ��������� ���� ������� � ������
    fireModeller.AddFireFromShape ShpObj

    MsgBox "������� ������ ��������� � ������� �������." & Chr(13) & Chr(13) & _
            "�������� ��������, ��� ��������� ������ ���������� �� ����� � ����� ���������� ��� �������! ����� �������� �����, ������� ��!"
End Sub
