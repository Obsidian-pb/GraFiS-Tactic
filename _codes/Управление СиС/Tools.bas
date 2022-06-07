Attribute VB_Name = "Tools"
Option Explicit

Public Function GetScaleAt200() As Double
'���������� ����������� ���������� ������� ������� �������� ������������ �������� 1:200
Dim v_Minor As Double
Dim v_Major As Double

    v_Minor = Application.ActivePage.PageSheet.Cells("PageScale").Result(visNumber)
    v_Major = Application.ActivePage.PageSheet.Cells("DrawingScale").Result(visNumber)
    GetScaleAt200 = (v_Major / v_Minor) / 200
End Function

Public Function GetGFSShapeTime(ByRef shp As Visio.Shape) As Double
    
    GetGFSShapeTime = cellVal(shp, "Prop.SetTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellVal(shp, "Prop.FormingTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function
    
    GetGFSShapeTime = cellVal(shp, "Prop.ArrivalTime", visDate)
    If GetGFSShapeTime > 0 Then Exit Function

GetGFSShapeTime = 0
End Function

Public Sub P_TryDeleteSmartTag(ByRef shp As Visio.Shape, stName As String, rowName As String)
'stName - �������� �����-����, rowName - �������� ������ �����-���� � ������ SmartTags
Dim i As Integer
Dim smartTagRowIndex As Integer
    
    '�������� ����� ����� ���� � ��������� ������� �� ����� ����� ���
    smartTagRowIndex = P_GetRowIndex(shp, rowName)
    If smartTagRowIndex >= 0 Then
        '��������� ��� ������ ������ Actions �� ������� ������� ������ �� ��������� ����� ���
        For i = 0 To shp.RowCount(visSectionAction) - 1
            '���� ���� ���� ���� - ������� �� ��������� �� ������ ����� ���
            If shp.CellsSRC(visSectionAction, i, visActionTagName).ResultStr(visUnitsString) = stName Then Exit Sub
        Next i
    
        shp.DeleteRow visSectionSmartTag, smartTagRowIndex
        '������� ��� �� � ������ ������������ �������
        On Error Resume Next
        shp.DeleteRow visSectionUser, shp.Cells("User.CurrentDocTime").row
    End If

End Sub

Public Function P_GetRowIndex(ByRef shp As Visio.Shape, cellName As String) As Integer
    On Error GoTo ex
    P_GetRowIndex = shp.Cells(cellName).row
Exit Function
ex:
    P_GetRowIndex = -1
End Function

'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub



