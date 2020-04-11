Attribute VB_Name = "Tools"
Option Explicit


'-----------------------------------------��������� ������ � ��������----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'��������� ������������� ����� �������� ��� ���� ��������� ����� ������ ����
Dim shp As Visio.Shape
    
    '���������� ��� ������ � ��������� � ���� ��������� ������ ����� ����� �� ������ - ����������� �� ����� ��������
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).formula = aB_Value
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


Public Sub HideMaster(ByVal masterName As String, ByVal visible As Integer)
'����� ��������/��������� ������ �� �����
Dim mstr As Visio.Master
Dim doc As Visio.Document

    Set doc = ThisDocument
    Set mstr = doc.Masters(masterName)
    mstr.Hidden = Not visible
    
    Set mstr = Nothing
    Set doc = Nothing
End Sub
'HideMaster "������ 2", 1 - ������
'HideMaster "������ 2", 0 - ������

Public Sub SeekBuilding(ShpObj As Visio.Shape)
'��������� ��������� ������� ������������� � ���������� ��� ������ ����������� ������� �������������
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim Col As Collection

    On Error GoTo EX
'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
        If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
            If OtherShape.Cells("User.IndexPers") = 135 And OtherShape.HitTest(x, y, 0.01) > 1 Then
                ShpObj.Cells("Prop.SO").FormulaU = _
                 "Sheet." & OtherShape.ID & "!Prop.SO"
            End If
        End If
    Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
'    SaveLog Err, "SeekFire", ShpObj.Name
End Sub
