VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CornerFillDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------����� ������ ����������� ���������� ����� ������� � ������ ������� (������� � ������������� �����)----------------------------
Const PI = 3.1415926535

Private matrix As c_Matrix

Public cellsForFill As Collection







'------------------------��������� ��������� ������---------------------------------
Public Sub Init(ByRef matrixRef As c_Matrix)
'�������������� �����
    Set matrix = matrixRef
    Set cellsForFill = New Collection
End Sub

Public Function testLine(ByVal startX As Long, ByVal startY As Long, ByVal angle As Double)
'���������, ��������� �� ����� ����������� �� ������������ ����� � ������������ ����� � ������� ����
Dim x As Integer
Dim y As Integer
Dim cs As Double
Dim sn As Double
Dim i As Integer
Dim j As Integer
Dim CellType As Byte
    
    '��������� ��������� ������
'    AddToCellCollection startX, startY
    
    '������� ���������� ������� � ����� ���� - ��� ����������� ���������� �����
    cs = Cos(angle * PI / 180)
    sn = Sin(angle * PI / 180)
    
    '��������������� ��������� ��� ������� �����, �� ������� ����, �������� ��� ������ �������, ������� ������� ��� �� ������� �����
    i = 0
    Do While CellType = 0
        x = startX + Int(cs * i)
        y = startY + Int(sn * i)
        
        CellType = GetKindOfCell(x, y)
        
        i = i + 1
        If i > 100 Then Exit Do
    Loop

    
    '����������� �����, ���� ��� �������� � ���� �������
'    '!!!��������, ��� ��������
'    If cellType = 1 Then
'        testLine = True
'        Exit Function
'    End If

    '������ ��� ������ � ��������� ������ ��� ������������ ������������
    If CellType = 1 Then
        For j = 0 To i - 2
            x = startX + Int(cs * j)
            y = startY + Int(sn * j)
            
            AddToCellCollection x, y
        Next j
    
        '���������� ������, ���� ����� ��������� �� ������� ������
        testLine = True
        Exit Function
    End If
    
    '���������� ����, ���� ����� �� ��������� �� ������� ������
    testLine = False
End Function


Private Function GetKindOfCell(ByVal xVal As Long, ByVal yVal As Long) As Byte
'���������� ��� ������ (0-���������, 1-����� �� ��������, 2-�����������-���������� ��������)
    On Error GoTo EX
    If matrix.IsCellMaxFire(xVal, yVal) Then
        GetKindOfCell = 1
        Exit Function
    End If
    If matrix.isWall(xVal, yVal) Then
        GetKindOfCell = 2
        Exit Function
    End If
    GetKindOfCell = 0
Exit Function
EX:
    GetKindOfCell = 2
End Function


Private Sub AddToCellCollection(ByVal xVal As Long, ByVal yVal As Long)
'��������� ������ � ���������
    Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = xVal
    cell.y = yVal
    
    On Error Resume Next
    cellsForFill.Add cell, xVal & ":" & yVal
End Sub


'Public Sub FillCells()
''�������� ��� ������ �� ���������
'Dim cell As c_Cell
'
'    For Each cell In cellsForFill
'        Cells(cell.y, cell.x) = 100
'    Next cell
'End Sub













