Attribute VB_Name = "m_Tools"
Option Explicit




Public Function RUp(ByVal val As Single, ByVal Modificator As Single) As Integer
'������� ���������� �������� ����� ������������ � ������� �������
Dim tmpVal As Single

On Error GoTo EX
    tmpVal = Int(val * Modificator / 10) * 10
    If tmpVal < 20 Then tmpVal = 20
    RUp = tmpVal
    
Exit Function
EX:
    RUp = Round(val)
End Function


Public Sub PS_GraphicsFix(ByRef shp As Visio.Shape)
'��������� ��������� ���� ������� ������������ ��������� ������ ����
Dim Time As String
Dim SqrExp As String
Dim TimeA() As String
Dim SqrExpA() As String
Dim i As Integer
Dim IndexPers As Integer

    On Error GoTo EX

    '---�������� ������� ������ ��� ��������� ������
    SqrExp = shp.Cells("Scratch.A1").ResultStr(visUnitsString)
    Time = shp.Cells("Scratch.B1").ResultStr(visUnitsString)
    StringToArray SqrExp, ";", SqrExpA()
    StringToArray Time, ";", TimeA()
    
    '---��� ������ �� ����� ������������ �������� � �����������
    IndexPers = shp.Cells("User.IndexPers")
    '---� ������ �������� ��������
    If IndexPers = 123 Or IndexPers = 124 Then
        For i = 0 To UBound(SqrExpA)
            shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(TimeA(i)) & "/User.TimeMax)*Width"
            shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(SqrExpA(i)) & "/User.FireMax)*Height"
        Next i
    End If
    '---� ������ �������� ��������
    If IndexPers = 125 Or IndexPers = 126 Then
        For i = 0 To UBound(SqrExpA)
            shp.Cells("Controls.Row_" & i + 1).FormulaU = "(" & str(TimeA(i)) & "/User.TimeMax)*Width"
            shp.Cells("Controls.Row_" & i + 1 & ".Y").FormulaU = "(" & str(SqrExpA(i)) & "/User.MaxExpense)*Height"
        Next i
    End If
    
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "PS_GraphicsFix"
End Sub

Public Sub StringToArray(ByVal str As String, ByVal Charackter As String, ByRef Arr() As String)
'��������� ����� �������� ������ �� ���������� ������� � ������������ ������
Dim i As Integer
Dim ValuesCount As Integer
Dim pos As Integer
    
    On Error GoTo EX
    
    '---���� ������ �� ������� �������� ��������, �� ��������� ��
    If Not Right(str, 1) = Charackter Then
        str = str & ";"
    End If
    
    '---���������� ���������� �������� � ������ � �������������� ������ ��������������� �������
    ValuesCount = 0
    For i = 0 To Len(str)
        If Mid(str, i + 1, 1) = Charackter Then ValuesCount = ValuesCount + 1
    Next i
    ReDim Arr(ValuesCount - 1)

    '--��������� ������
    For i = 0 To ValuesCount - 1
        pos = InStr(1, str, Charackter, vbTextCompare)
        Arr(i) = Left(str, pos - 1)
        str = Right(str, Len(str) - pos)
    Next i
    
Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "StringToArray"
End Sub

Public Sub PS_GetTotelExpense(ByRef shp As Visio.Shape)
'����� ������������� ��� ������ ������� ������ �������� ������ ������� ���� �� ��� �����
    shp.Cells("Prop.TotalExpence").FormulaForceU = "Guard(" & Int(pf_GetTotalExpence(shp)) & ")"
End Sub

Private Function pf_GetTotalExpence(ByRef shp As Visio.Shape) As Double
'������� ���������� ��������� ������ ���� ��� ������ �������, � ������
Dim vL_SecondsCount As Long      '������������ �����
Dim vL_SecondEpenceValue As Long '������������ ��������� ������
Dim vD_BlockDuration As Double   '����������������� ����� � ��������
Dim vD_BlockExpence As Double    '��������� ������ �����
Dim vD_TotalExpence As Double    '����� ������ �����(�)
Dim i As Integer

On Error GoTo Tail

    vL_SecondsCount = shp.Cells("User.TimeMax").Result(visNumber) * 60
    vL_SecondEpenceValue = shp.Cells("User.MaxExpense").Result(visNumber)
    
    For i = 0 To shp.RowCount(visSectionControls) - 1
        vD_BlockDuration = (shp.CellsSRC(visSectionFirstComponent, i * 2 + 4, 0).Result(visMeters) - _
                            shp.CellsSRC(visSectionFirstComponent, i * 2 + 3, 0).Result(visMeters)) / shp.Cells("Width").Result(visMeters) * vL_SecondsCount
        vD_BlockExpence = shp.CellsSRC(visSectionControls, i, 1).Result(visMeters) / _
                            shp.Cells("Height").Result(visMeters) * vL_SecondEpenceValue
        vD_TotalExpence = vD_TotalExpence + vD_BlockDuration * vD_BlockExpence
    Next i

pf_GetTotalExpence = vD_TotalExpence
Exit Function
Tail:
    pf_GetTotalExpence = 0
End Function

'-----------------------------------------��������� ������ � ��������----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'��������� ������������� ����� �������� ��� ���� ��������� ����� ������ ����
Dim shp As Visio.Shape
    
    '���������� ��� ������ � ��������� � ���� ��������� ������ ����� ����� �� ������ - ����������� �� ����� ��������
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).Formula = aB_Value
        End If
    Next shp
    
End Sub


'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub
