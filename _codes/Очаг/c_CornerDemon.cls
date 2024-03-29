VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CornerDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------����� ������ ����������� ����� ����------------------------------------------------
Const PI = 3.1415926535

Private x As Long
Private y As Long
Private startX As Long
Private startY As Long
Private direction As Directions

Private mainAngle As Double
Private perpAngle As Double

Private matrix As c_Matrix
Public lineFillDemon As c_CornerFillDemon




'-----------------��������� ������ � ������� �����----------------------------------------
Public Sub Init(ByRef matrixRef As c_Matrix, ByVal xVal As Long, ByVal yVal As Long)
'������������� ������
    direction = s       '����� �� �����
    
    '�������� �������
    Set matrix = matrixRef
    
    '���������� ������-����������� �����
    Set lineFillDemon = New c_CornerFillDemon
    lineFillDemon.Init matrix
    
    '������ ��������� �������
    SetStartPosition xVal, yVal
End Sub

Public Sub SetStartPosition(ByVal xVal As Long, ByVal yVal As Long)
'������ ��������� ���������� ������
    startX = xVal
    startY = yVal
    SetPosition xVal, yVal
End Sub

Private Sub SetPosition(ByVal xVal As Long, ByVal yVal As Long)
'������ ���������� ������
    x = xVal
    y = yVal
End Sub

Public Sub SetDirection(ByVal dirct As Directions)
'������ ����������� �������� ������
    direction = dirct
End Sub

Public Sub RunDemon()
'�������� ������
Dim i As Integer
    
    '��������� ������� ������ ��� ����������� ���� ����������������� �����
    Dim angleDemon As c_AngleDemon
    Set angleDemon = New c_AngleDemon
    angleDemon.Init matrix, x, y, direction
    angleDemon.RunDemon
    
    mainAngle = angleDemon.mainAngle
    perpAngle = angleDemon.perpAngle
    
    
    
    '����� ��������� �� ���� ������� ����� ����� � ������ ��� ��������� ������������� � ���. ���� ����� �� ����� ��������� ���������� ��������, ���� �� ����� ������� �� ����� �� �������������� ������� ������, ����������� �������
    i = 0
    Do While NextCell
        i = i + 1
        '��������� �� ������� �� ������ ������ �� ������� ����������� ����
        If Not isPossiblePlace Then Exit Do
        
        '���������, �� ������� �� ������ ������ �� ������� �������
        If matrix.IsCellOutOfArea(x, y) Then Exit Sub
        
        '��������� ������ �������������� ����� �� �������� ���������
        If lineFillDemon.testLine(x, y, perpAngle) = False Then Exit Do
        
        '������� - ���������� �����, ���� ���-�� ����� �� ���
        If i > 100 Then
            Exit Do
        End If
    Loop
       
End Sub




'-------------------------------------------���������� ��������� ������----------------------------------------
Private Function NextCell() As Boolean
'�������� ��������� ������

    '���� ����� ���� �� �������� - ���� �����������
    If direction = s Then
        direction = FindDirection
        
        '���� ����� ����������� �������� �� ������� - �������
        If direction = s Then
            NextCell = False
            Exit Function
        End If
        
        '���������� ������
        MoveDemon
        NextCell = True
        Exit Function
    Else
        '���� ������ � ����������� ����� �� ���� ����������
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(0)) Then
            direction = NewDirection(0)
            '���������� ������
            MoveDemon
            NextCell = True
            Exit Function
        End If
        '���� ������ � ����������� ����� �� ���� ����������
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(-2)) Then
            direction = NewDirection(-2)
            '���������� ������
            MoveDemon
            NextCell = True
            Exit Function
        End If
        '���� ������ � ����������� ������ �� ���� ����������
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(2)) Then
            direction = NewDirection(2)
            '���������� ������
            MoveDemon
            NextCell = True
            Exit Function
        End If
    End If
    
    '���������� ���� �� ������
    direction = s
    MoveDemon
    NextCell = False
    
End Function

Private Sub MoveDemon()
'���������� ������ � ������� �����������
    Select Case direction
        Case Is = l
            x = x - 1
        Case Is = lu
            y = y + 1
            x = x - 1
        Case Is = u
            y = y + 1
        Case Is = ru
            y = y + 1
            x = x + 1
        Case Is = r
            x = x + 1
        Case Is = rd
            y = y - 1
            x = x + 1
        Case Is = d
            y = y - 1
        Case Is = ld
            y = y - 1
            x = x - 1
        Case Else
'            Debug.Print "�����!!! x=" & x & ", y=" & y
    End Select
End Sub

Private Function IsCellNearTheWallAndNotFireByDirection(ByVal direct As Directions) As Boolean
    Select Case direct
        Case Is = l
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y)
        Case Is = lu
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y + 1)
        Case Is = u
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y + 1)
        Case Is = ru
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y + 1)
        Case Is = r
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y)
        Case Is = rd
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y - 1)
        Case Is = d
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y - 1)
        Case Is = ld
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y - 1)
    End Select
End Function

Private Function IsCellNearTheWallAndNotFire(ByVal xVal As Long, ByVal yVal As Long) As Boolean
'�������� ��������� �� ������ ����� �� ������.  � ��� ���� �� ��������� ����� � ���� �� �������� ������
    '���������, �� ������� �� ���������� ������ �� ������� �������
    If x < 0 Or y < 0 Or x > matrix.DimensionX Or y > matrix.DimensionY Then Exit Function
    
    '��������� ���� ������
    If matrix.IsCellMaxFire(xVal, yVal) Or matrix.isWall(xVal, yVal) Then
        IsCellNearTheWallAndNotFire = False
        Exit Function
    End If
    
    '�� ����������
    If matrix.isWall(xVal, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal - 1, yVal) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal + 1, yVal) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    '�� ���������
    If matrix.isWall(xVal - 1, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal - 1, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal + 1, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.isWall(xVal + 1, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    
    
IsCellNearTheWallAndNotFire = False
End Function

Private Function NewDirection(ByVal directionDifferent As Integer) As Directions
'�������� �������� ����������� � ����������� � ������������ �������������
Dim tmpVal As Integer
    
    tmpVal = direction + directionDifferent
    
    If tmpVal < 1 Then tmpVal = tmpVal + 8
    If tmpVal > 8 Then tmpVal = tmpVal - 8
    
NewDirection = tmpVal
End Function

Private Function FindDirection() As Directions
    
    '�� ���������
    If IsCellNearTheWallAndNotFire(x - 1, y - 1) Then
        FindDirection = ld
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y - 1) Then
        FindDirection = rd
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y + 1) Then
        FindDirection = lu
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y + 1) Then
        FindDirection = ru
        Exit Function
    End If
    '�� ����������
    If IsCellNearTheWallAndNotFire(x, y - 1) Then
        FindDirection = d
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x, y + 1) Then
        FindDirection = u
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y) Then
        FindDirection = l
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y) Then
        FindDirection = r
        Exit Function
    End If
End Function

Private Function GetAngle() As Double
'���������� ���� ����� ��������� �������� � ������� ���������� ������
Dim xDiff As Long
Dim yDiff As Long

    On Error Resume Next

    xDiff = x - startX
    yDiff = y - startY


    If xDiff = 0 Then
        If y > startY Then
            GetAngle = 90
        Else
            GetAngle = -90
        End If
    Else
        '������������ ��������
        '1-� ��������
        If xDiff > 0 And yDiff > 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI
        End If
        '2-� ��������
        If xDiff < 0 And yDiff > 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI + 180
        End If
        '3-� ��������
        If xDiff < 0 And yDiff < 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI + 180
        End If
        '4-� ��������
        If xDiff > 0 And yDiff < 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI
        End If
        
    End If

End Function

Private Function isPossiblePlace() As Boolean
'������� ���������� ������, ���� ������ ���������� ��������� ��� ������ � ������ ����� � ������ �����
Dim possibleDiff As Double
Dim tolerance As Single
Dim dist As Double
Dim currentAngle As Double
    
    '���������� ���������� ������ ���������� ����
    dist = Sqr((x - startX) ^ 2 + (y - startY) ^ 2)
    If dist = 0 Then
        isPossiblePlace = True
        Exit Function
    End If
    
    tolerance = 1.01
    possibleDiff = Atn(1 / dist) * (180 / PI) * tolerance
    
    '�������� ������� ����
    currentAngle = GetAngle
    
    '���������, ����� �� ���� � ������� ������ � ���������� ��������
    If currentAngle < mainAngle + possibleDiff And currentAngle > mainAngle - possibleDiff Then
        isPossiblePlace = True
    Else
        isPossiblePlace = False
    End If
    
End Function
