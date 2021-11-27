Attribute VB_Name = "FightingActionsAnimation"
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
 
Private Declare Sub GetSystemTime Lib "kernel32" (st As SYSTEMTIME)

Const interval = 5
Const secondsInTurn = 10     '��������� ������� �������� �� ��� - ���� 1
Const �hangeVal = 16.6         '�������� ��������� - ���� 5

Private turns As Long
Private startDateTime As Date
Private endDateTime As Date
Private CellsList() As String


Private Function milSec(ByRef tt As SYSTEMTIME) As Long
Dim s As Long
Dim m As Long
    
    s = CLng(tt.wSecond) * 1000
    m = CLng(tt.wMinute) * 60000

    l = tt.wMilliseconds + s + m
'    Debug.Print tt.wMilliseconds + tt.wSecond * 1000  '+ tt.wMinute * 60000
    milSec = l   'tt.wMilliseconds + tt.wSecond * 1000 '+ tt.wMinute * 60000
End Function

Private Sub FillCellsList()
ReDim CellsList(9) As String
    CellsList(0) = "Prop.ArrivalTime"
    CellsList(1) = "Prop.SetTime"
    CellsList(2) = "Prop.LineTime"
    CellsList(3) = "Prop.SquareTime"
    CellsList(4) = "Prop.FireTime"
    CellsList(5) = "Prop.UTPCreationTime"
    CellsList(6) = "Prop.FormingTime"
    CellsList(7) = "Prop.StabCreationTime"
    CellsList(8) = "Prop.ApearanceTime"              '����� ��������� ������������� �������
    CellsList(9) = "Prop.FindTime"
End Sub



Public Sub runMutation()
Dim shp As Visio.Shape
Dim tt As SYSTEMTIME
Dim i As Long
Dim maxI As Long

Dim begin As SYSTEMTIME
Dim cur As SYSTEMTIME
Dim curDateTime As Date
    
    '��������� ������ �������� ����� ��� �������� �������
    FillCellsList
    
    '������������� �������������� ��� ���� �����
    ShapeTranspSetForAllShapes 100
    
    '�������� ��������� ����
'    startDateTime = Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    startDateTime = CDate("15.04.2021 12:25:00")
    endDateTime = CDate("15.04.2021 12:55:00")
    turns = 0
    
    
    GetSystemTime begin
    GetSystemTime cur
    
'    Do While i < 1000
    Do While curDateTime < endDateTime
        If milSec(cur) > milSec(begin) + interval Then
            'Action:
            turns = turns + 1
            
            curDateTime = DateAdd("s", turns * secondsInTurn, startDateTime)
            CheckAllShapesTime curDateTime
            
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Formula = "DATETIME(" & CDbl(curDateTime) & ")"
            Application.CommandBars("������").Controls("�����").Text = TimeValue(curDateTime)
            Application.CommandBars("������").Controls("�����").SetFocus

            Debug.Print turns & ", " & curDateTime
            
            DoEvents
            
            i = i + 1
            begin = cur
        End If
        
        GetSystemTime cur
    Loop

End Sub



Public Sub CheckAllShapesTime(ByVal curTime As Date)
'��������� ����� ���� ����� � ���� ��� ������ ��������, �������� ���
Dim shp As Visio.Shape
    
    For Each shp In Application.ActivePage.Shapes
        If IsTimeLeft(shp, curTime) Then
            ShapeTranspChange shp, -1 * �hangeVal, True
        End If
'        If ShpIsInvisible(shp) Then
'            ShapeTranspChange shp, 100, True
'        End If
    Next shp
End Sub

Private Function IsTimeLeft(ByRef shp As Visio.Shape, ByVal curTime As Date) As Boolean
Dim i As Byte
    
    For i = 0 To UBound(CellsList)
        If IsTimeLeftForCell(shp, curTime, CellsList(i)) Then
            IsTimeLeft = True
            Exit Function
        End If
    Next i
    
IsTimeLeft = False
'    IsTimeLeft = IsTimeLeftForCell(shp, curTime, "Prop.ArrivalTime") + _
'                 IsTimeLeftForCell(shp, curTime, "Prop.SetTime") + _
'                 IsTimeLeftForCell(shp, curTime, "Prop.LineTime") + _
'                 IsTimeLeftForCell(shp, curTime, "Prop.SquareTime") + _
'                 IsTimeLeftForCell(shp, curTime, "Prop.FireTime")
End Function

Private Function IsTimeLeftForCell(ByRef shp As Visio.Shape, ByVal curTime As Date, ByVal cellName As String) As Boolean
Dim shpTime As Date
    
    On Error GoTo ex
    
    IsTimeLeftForCell = False
    
'    shpTime = shp.Cells(cellName).Result(visDate)
    shpTime = shp.Cells(cellName).ResultStr(visUnitsString)
'    If shpTime < curTime Then  'And shpTime > DateAdd("s", -20, curTime) Then
'    If shpTime < curTime And DateAdd("s", 20, shpTime) > curTime Then
    If shpTime < curTime And DateAdd("s", (100 / �hangeVal) * secondsInTurn, shpTime) > curTime Then
        IsTimeLeftForCell = True
    End If
    If shpTime < curTime Then
        If DateAdd("s", (100 / �hangeVal) * secondsInTurn, shpTime) > curTime Then
            IsTimeLeftForCell = True
        Else
            If ShpIsInvisible(shp) Then
                IsTimeLeftForCell = True
                ShapeTranspSet shp, 0
            End If
        End If
    End If
    
    
Exit Function
ex:
    IsTimeLeftForCell = False
End Function




Private Sub ShapeTranspChange(ByRef shp As Visio.Shape, ByVal val As Integer, Optional ByVal grafisCheckNeed = True)
Dim curVal As Integer
Dim shpChild As Visio.Shape
    
    '���� ����� ��������, �� ��������� �������� �� ������ ��� ������������
    If grafisCheckNeed = True Then
        If Not IsTimedGrafisShape(shp) Then Exit Sub
    End If
    
    '���������� ��� ������
    If shp.Shapes.Count > 0 Then
        For Each shpChild In shp.Shapes
            ShapeTranspChange shpChild, val, False
        Next shpChild
    End If
    
    On Error Resume Next
    
    '�������� �������� ��-���������
    curVal = shp.Cells("LineColorTrans").Result(visPercent)

    If curVal < 0 Then Exit Sub
    '�������� ��������
    curVal = curVal + val
    
    If curVal >= 0 Then
        shp.Cells("LineColorTrans").Formula = curVal & "%"
        shp.Cells("FillForegndTrans").Formula = curVal & "%"
        shp.Cells("FillBkgndTrans").Formula = curVal & "%"
        shp.Cells("Char.ColorTrans").Formula = curVal & "%"
    ElseIf curVal < 0 Then
        shp.Cells("LineColorTrans").Formula = "0%"
        shp.Cells("FillForegndTrans").Formula = "0%"
        shp.Cells("FillBkgndTrans").Formula = "0%"
        shp.Cells("Char.ColorTrans").Formula = "0%"
    End If
    
    
End Sub



Private Sub ShapeTranspSetForAllShapes(ByVal val As Integer)
Dim shpChild As Visio.Shape
    
    For Each shpChild In Application.ActivePage.Shapes
        ShapeTranspSet shpChild, val, True
    Next shpChild
End Sub

Private Sub ShapeTranspSet(ByRef shp As Visio.Shape, ByVal val As Integer, Optional ByVal grafisCheckNeed = True)
Dim curVal As Integer
Dim shpChild As Visio.Shape
    
    '���� ����� ��������, �� ��������� �������� �� ������ ��� ������������
    If grafisCheckNeed = True Then
        If Not IsTimedGrafisShape(shp) Then Exit Sub
    End If
    
    '���������� ��� ������
    If shp.Shapes.Count > 0 Then
        For Each shpChild In shp.Shapes
            ShapeTranspSet shpChild, val, False
        Next shpChild
    End If

    On Error Resume Next
    
    shp.Cells("LineColorTrans").Formula = val & "%"
    shp.Cells("FillForegndTrans").Formula = val & "%"
    shp.Cells("FillBkgndTrans").Formula = val & "%"
    shp.Cells("Char.ColorTrans").Formula = val & "%"
    
End Sub

Private Function IsTimedGrafisShape(ByRef shp As Visio.Shape) As Boolean
'��������� �������� �� ������ ��� ������������
Dim i As Byte
    
    If shp.CellExists("User.IndexPers", 0) = False Then
        IsTimedGrafisShape = False
        Exit Function
    End If
    
    For i = 0 To UBound(CellsList)
        If shp.CellExists(CellsList(i), 0) Then
            IsTimedGrafisShape = True
            Exit Function
        End If
        
    Next i

IsTimedGrafisShape = False
End Function

Private Function ShpIsInvisible(ByRef shp As Visio.Shape) As Boolean
    ShpIsInvisible = shp.Cells("LineColorTrans").Result(visNumber) > 0  ' = 1
End Function

'Public Sub ToolAddProp()
'Dim shp As Visio.Shape
'Dim rowI As Integer
'
'    On Error Resume Next
'
''    startDateTime = Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
'
'    For Each shp In Application.ActiveWindow.Selection
'        rowI = shp.AddNamedRow(visSectionProp, "ApearanceTime", visTagDefault)
'        shp.CellsSRC(visSectionProp, rowI, visCustPropsType).FormulaU = 5
'    Next shp
'End Sub
