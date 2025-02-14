VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_MatrixAreaBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------����� ��� ��������� ������� ��������� ����������� ��� �������----------------------------------------------------------------------------------
Option Explicit


Private grain As Integer        '������ ����� � ��
Const mmInInch As Double = 25.4

Private form As Object

Private bordreRect As Visio.Shape
Private bordreOpen As Visio.Shape
Private points As Collection

Public matrix As Variant

Const radiussModificator As Double = 1.4142



Public Function SetForm(ByRef frm As Object) As C_MatrixAreaBuilder
'��������� ������ �� ������� ����� � ������� ����� ���������� ����� � �������� ��������� �����
    Set form = frm
    
Set SetForm = Me
End Function


Public Function FillArea(ByRef matrix_a As Variant, ByRef bordreRect_a As Visio.Shape, ByVal grain_a As Integer) As Variant
'������� ������� � ���������� ������� ���������� ������������ (����, ������� ����� ����������� � �������)
'grain - ������ ����� � ��
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer
    
    grain = grain_a
    matrix = matrix_a
    Set bordreRect = bordreRect_a
    
    '---������� ������-������
    NewScaner
    
    '---���������� ���������
    CheckEnvironment


    '---���������� ��������� - ������� ����������� ������� � ��������� ������������
    FillArea = matrix
    
    '������� ������
    bordreOpen.Delete
    

End Function

Private Sub CheckEnvironment()
'��������� ���������� ������������ ������������
Dim startRectSizeHalf As Long
    
    '���������� ����������� �����
    '---���������� ������ ���������� �������� ��� ������ ����
        startRectSizeHalf = GetMaximumRectSize / 2   '�������� ������ ��� � ����� ������ � ������ � ���������� ����� �����
'        '������ ��� ������� ����� �������� ������� �������
        form.SetMatrixSize (startRectSizeHalf * 2) ^ 2
    '---��������� ����������� ������
        CheckRect startRectSizeHalf, startRectSizeHalf, startRectSizeHalf
End Sub

Private Function IsOpenSpace(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ����, � ��������, ���� �������������
    IsOpenSpace = bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function

Private Function IsTotalWall(x As Long, y As Long, tolerance As Double) As Boolean
'������� ���������� ������, ���� ��������� ����� �� ������������� � ������ ���������� ������������, � ��������, ���� �������������
    IsTotalWall = bordreOpen.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function



'------------------����������� ����������� �������------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'����� radiuss - � ������!!!!
Dim radiussHalf As Integer

    On Error GoTo EX
    Dim shp As Visio.Shape
    
'1 ��������� �����������, ��� � ������� ��� ��������� ������� (����� - ��)
    '---���� ���� ���, ��������� ��� ���� ������ 0
        If IsOpenSpace(x, y, radiuss * radiussModificator * grain) Then
            form.AddCheckedSize (radiuss * 2) ^ 2, 1
'            form.Refresh
'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                    (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'            shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
            DoEvents

            Exit Sub
        End If
        
'2 ���� ����� ����, ��������� ������� ��������
'3 ��������� �����������, ��� � ������� ��� ��������� ������������
    '---���� ����� ����, ��������� ����� ��������
        If IsTotalWall(x, y, radiuss * radiussModificator * grain) Then
        '���� ��� ���� � ������� ��� ��������� ������������, ������ ��� ���� �������� ����� �������� ����� (1)
            SetToAllCells x, y, radiuss, CellState.csOpenSpace
            form.AddCheckedSize (radiuss * 2) ^ 2, 1
            
'                Dim shp As Visio.Shape
'                Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                        (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
''
                DoEvents
        Else
            If radiuss = 1 Then
                form.AddCheckedSize 1, 1
                
                '��������� ��� �������� ������, �.�. ��� ������� = 1 � ������ ������ ������ ���������
                '1 - ����� �����
                    CheckSingleCell x, y
                '2 - ����� ������
                    CheckSingleCell x, y - 1
                '3 - ���� ������
                    CheckSingleCell x - 1, y
                '4 - ���� �����
                    CheckSingleCell x - 1, y - 1
                
                Exit Sub
            Else
                '����� ���������� �������� �������� ������� �������� ���������
                radiussHalf = radiuss * 0.5
                '1 - ����� �����
                    CheckRect x - radiussHalf, y + radiussHalf, radiussHalf
                '2 - ����� ������
                    CheckRect x + radiussHalf, y + radiussHalf, radiussHalf
                '3 - ���� ������
                    CheckRect x + radiussHalf, y - radiussHalf, radiussHalf
                '4 - ���� �����
                    CheckRect x - radiussHalf, y - radiussHalf, radiussHalf
            End If
        End If
Exit Sub
EX:
    Debug.Print "x=" & x & ", y=" & y & ", radiuss=" & radiuss
End Sub

Private Sub CheckSingleCell(ByVal x As Integer, ByVal y As Integer, Optional tolerance As Double = 0)
    If bordreOpen.HitTest((x + 0.5) * grain / mmInInch, (y + 0.5) * grain / mmInInch, tolerance / mmInInch) = 0 Then
        matrix(x, y) = CellState.csOpenSpace
    End If
End Sub



Private Sub SetToAllCells(ByVal x As Long, ByVal y As Long, ByVal size As Long, ByVal value As CellState)
'������������� ��� ���� ������ ���������� ��������� size � ������� � ����� x,y �������� value
Dim i As Long
Dim j As Long

    For i = x - size To x + size
        For j = y - size To y + size
            matrix(i, j) = value
        Next j
    Next i
End Sub

Private Function GetMaximumRectSize() As Long
'���������� ����������� ��������� ������� ��������� ����������� ����� � ��� ���� ���������� ��� ������� ����� ��������� ��������� �� 2
Dim grains As Long
Dim pageWidth As Double
Dim pageHeight As Double

    grains = 1
    
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)
    
    Do While grains * grain < pageWidth Or grains * grain < pageHeight
        grains = grains * 2
    Loop
    
GetMaximumRectSize = grains
End Function

'------------------����������� ����������� �������------------------------------------------------------






'------------------������� �������--------------------------
Private Sub NewScaner()
Dim shp1 As Visio.Shape
Dim shp2 As Visio.Shape

    Set points = New Collection
    
    Set shp1 = Application.ActivePage.DrawRectangle(0, 0, Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches), Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches))
    Set shp2 = Application.ActivePage.Drop(bordreRect, bordreRect.Cells("PinX").Result(visInches), bordreRect.Cells("PinY").Result(visInches))
    
    '---��������� � ��������� ������������� �������� � ����� ����
    Application.ActiveWindow.Selection.DeselectAll
    Application.ActiveWindow.Select shp2, visSelect
    Application.ActiveWindow.Select shp1, visSelect
    '---�������� �� �������������� ����� ���� ����� ����
    Application.ActiveWindow.Selection.Combine
    
    Set bordreOpen = Application.ActiveWindow.Selection(1)
    
End Sub

Private Function Scan(ByVal x As Double, ByVal y As Double, Optional ByVal size As Double = 1) As Byte
'0 - scanner �� ������ � bordreRect, 1 - �������������, 2 - ��������� ������������
Dim val As Byte

    val = ScanPoint(x, y) + ScanPoint(x, y - size) + ScanPoint(x - size, y) + ScanPoint(x - size, y - size)
    If val = 0 Then
        Scan = 0
    ElseIf val = 4 Then
        Scan = 2
    Else
        Scan = 1
    End If
End Function
Private Function ScanSingle(ByVal x As Double, ByVal y As Double, Optional ByVal size As Double = 1) As Byte
'0 - scanner �� ������ � bordreRect, 1 - �������������, 2 - ��������� ������
Dim val As Byte
Dim idx As String

    ScanSingle = bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, size)
    
    idx = XYtoIndex(x, y)
    If ScanSingle = 0 Then
        points.Add False, idx
    ElseIf val = 2 Then
        points.Add True, idx
    Else
        points.Add False, idx
    End If
End Function

Private Function ScanPoint(ByVal x As Long, ByVal y As Long) As Boolean
'���������� True ���� ����� ������ ��������� ����, ����� False
Dim idx As String

    On Error GoTo EX
    
    idx = XYtoIndex(x, y)
    ScanPoint = points.item(idx)
    
Exit Function
EX:
    If bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, 0) = 2 Then
        ScanPoint = True
    Else
        ScanPoint = False
    End If
    points.Add ScanPoint, idx
End Function

Private Function XYtoIndex(ByVal x As Long, ByVal y As Long) As String
    XYtoIndex = str(x) & ":" & str(y)
End Function
