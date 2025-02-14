VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_PerimeterDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conturPoints As Collection                                  '��������� ����� �������

Public grain As Integer
Const mmInInch As Double = 25.4

'--------------------------------����� ����������� ��������� (������� ������)-------------------------------------------



Private Sub Class_Initialize()
    Set conturPoints = New Collection
    Set perimeterCells = New Collection
End Sub


Public Function setGrain(ByVal gr As Integer) As c_PerimeterDemon
    grain = gr
    
    Set setGrain = Me
End Function


Public Sub RunDemon(ByRef perimeterCellsCollection As Collection)
    
    '������ ������������������ ������
    BuildConturLine perimeterCellsCollection

End Sub



Private Sub BuildConturLine(ByRef perimeterCellsCollection As Collection)
'��������� ������ ������ ���������
Dim cell As c_Cell
Dim shp As Visio.Shape
    
'    ClearLayer "������������"
    
    For Each cell In perimeterCellsCollection
        BuildConturElement cell
        
'        Set shp = Application.ActivePage.DrawRectangle((cell.x * grain) / mmInInch, (cell.y * grain) / mmInInch, _
'                        ((cell.x + 1) * grain) / mmInInch, ((cell.y + 1) * grain) / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("������������")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    Next cell
End Sub

Private Sub BuildConturElement(ByRef cell As c_Cell)

Dim neighbour(0 To 3) As Boolean
Dim pnt As c_Point
Dim pnt2 As c_Point
    
    '���������� ������������� ������� ������
    neighbour(0) = cell.NCell(l).IsCellMaxFire
    neighbour(1) = cell.NCell(u).IsCellMaxFire
    neighbour(2) = cell.NCell(r).IsCellMaxFire
    neighbour(3) = cell.NCell(d).IsCellMaxFire
    
    '��������� ��� ������ ������������� ������� ������ � ��� ������ �� ��� ��� ������������� ��������� �����
    '---��������� ������ �� �������� ������, �� ��, ����� �� ��� �� ���������
    If Not neighbour(0) Then
        Set pnt = GetConturPoint(cell.x, cell.y)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x, cell.y + 1)
    End If
    If Not neighbour(1) Then
        Set pnt = GetConturPoint(cell.x, cell.y + 1)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x + 1, cell.y + 1)
    End If
    If Not neighbour(2) Then
        Set pnt = GetConturPoint(cell.x + 1, cell.y + 1)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x + 1, cell.y)
    End If
    If Not neighbour(3) Then
        Set pnt = GetConturPoint(cell.x + 1, cell.y)
        AddPointToCollection conturPoints, pnt
        pnt.SetNextPoint GetConturPoint(cell.x, cell.y)
    End If

End Sub


Private Sub AddPointToCollection(ByRef thisColl As Collection, ByVal pnt As c_Point)
'��������� � ��������� ��������� ��������� �����
    On Error Resume Next
    thisColl.Add pnt, pnt.x & ":" & pnt.y
End Sub
Private Sub RemovePointFromCollection(ByRef thisColl As Collection, ByVal pnt As c_Point)
'������� ��������� ����� �� ��������� ���������
    On Error Resume Next
    thisColl.Remove pnt.x & ":" & pnt.y
End Sub


Private Function GetConturPoint(ByVal x As Long, ByVal y As Long) As c_Point
Dim pnt As c_Point
    On Error GoTo EX
    Set GetConturPoint = conturPoints.item(x & ":" & y)

Exit Function
EX:
    Set pnt = New c_Point
    pnt.SetData x, y
    AddPointToCollection conturPoints, pnt
    Set GetConturPoint = pnt
End Function




'-----------���������� ��������� �������------------
Public Function DrawSquare() As Visio.Shape
'������������ ������� ��������� ��� ����� ���������� DrawPoiyLine
Dim pnt As c_Point
Dim pntStart As c_Point

Dim shp As Visio.Shape
Dim perimterPnts() As Double
Dim i As Long

    On Error GoTo EX

'#� PerimeterCellsFix
    '���������� ������������������ �����
    PerimeterCellsFix
'    PerimeterCellsFix2  '����������
    
    
    Dim coll As Collection
    For Each coll In conturPoints
        If coll.Count > 1 Then
            i = 0
            ReDim perimterPnts(coll.Count * 2 + 1)
        
            '�������������� ������ � ������������ ���������
            For Each pnt In coll
                perimterPnts(i) = (pnt.x * grain) / mmInInch
                i = i + 1
                perimterPnts(i) = (pnt.y * grain) / mmInInch
                i = i + 1
                
                If i > 100000 Then
                    Debug.Print "������������ � ����� ��������� ���������"    ', vbCritical, "��������� ���������"
                    Stop
                    
                    Set DrawSquare = Nothing
                    Exit Function
                End If
            Next pnt
            '---�������������� �����, ���������� ������
            Set pnt = coll(1)
            perimterPnts(i) = (pnt.x * grain) / mmInInch
            i = i + 1
            perimterPnts(i) = (pnt.y * grain) / mmInInch
            
            '������ ������ �� ������� ����� � ������� �� ������������ ��������
            Set shp = ActivePage.DrawPolyline(perimterPnts, 0)
                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
                shp.Cells("LinePattern").FormulaU = "0"
                shp.Cells("FillForegnd").FormulaU = "2"
                shp.Cells("FillForegndTrans").FormulaU = "50%"
                shp.Cells("FillBkgndTrans").FormulaU = "50%"
        End If
    Next coll
    

Set DrawSquare = shp
EX:
    Set DrawSquare = Nothing
End Function

Private Sub PerimeterCellsFix()
'���������� ������������������ �����
Dim pnt As c_Point
Dim tmpPnt As c_Point
Dim curDir As Directions
Dim newPerimeters As Collection
Dim newPerimeterPoints As Collection
Dim i As Long

    Set newPerimeters = New Collection

    Do While conturPoints.Count > 0
        Set newPerimeterPoints = New Collection
        curDir = s
        
        Set pnt = conturPoints(1)
        Set tmpPnt = New c_Point
        tmpPnt.SetData pnt.x, pnt.y
        
        i = 0
        Do
            Set pnt = pnt.GetNextPoint
            '��������� � ����� ����������� ������ ��������� ����� �������
            If Not pnt.Dir = curDir Then    '���� ����������� ���������� - ��������� ����� ����� � ��������� � ������ ������� �����������
                AddPointToCollection newPerimeterPoints, pnt
    
                curDir = pnt.Dir
            End If
            
            RemovePointFromCollection conturPoints, pnt
            
            i = i + 1
            If i > 100000 Then
                Debug.Print "������������ � ����� ��������� ���������"    ', vbCritical, "��������� ���������"
                Stop
                Exit Sub
            End If
        
        Loop Until pnt.isEqual(tmpPnt)
        
        newPerimeters.Add newPerimeterPoints
    Loop
    

    '��������� ���������
    Set conturPoints = newPerimeters
End Sub



'Private Sub PerimeterCellsFix2()
''������, ��� ��� ������������
''���������� ������������������ ����� - ��������� ����
'Dim curPnt As c_Point
'
'Dim coll As Collection
'Dim newMainColl As Collection
'Dim newColl As Collection
'Dim i As Long
'Dim j As Long
'
'    On Error GoTo EX
'
'    Set newMainColl = New Collection
'
'    j = 0
'    For Each coll In conturPoints
'        i = 0
'        j = j + 1
'        Set newColl = New Collection
'
'        '�������������� ������ � ������������ ���������
'        For i = 1 To coll.Count
'            If Not IsOdd(i) Then
'                AddPointToCollection newColl, coll(i)  'curPnt
'            End If
'
'        Next i
'        newMainColl.Add newColl
'
'    Next coll
'
'    Set conturPoints = newMainColl
'Exit Sub
'EX:
'    Stop
'End Sub

Private Function GetCornerLen(ByRef pnt1 As c_Point, ByRef pnt2 As c_Point, ByRef pnt3 As c_Point) As Integer
'���������� ��������� ����� �������� ����� ����� �������
Dim tmpVal As Integer
    
    If Abs(pnt1.x - pnt2.x) = 1 Then tmpVal = tmpVal + 1
    If Abs(pnt1.y - pnt2.y) = 1 Then tmpVal = tmpVal + 1
    If Abs(pnt2.x - pnt3.x) = 1 Then tmpVal = tmpVal + 1
    If Abs(pnt2.y - pnt3.y) = 1 Then tmpVal = tmpVal + 1

GetCornerLen = tmpVal
End Function

Private Function isOdd(ByVal value As Integer) As Boolean
    isOdd = value / 2 = Int(value / 2)
End Function
