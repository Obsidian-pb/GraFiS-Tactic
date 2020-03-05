Attribute VB_Name = "m_Tools"
'-----------------------------------------������ ���������������� �������----------------------------------------------
Option Explicit



Public Function PF_RoundUp(afs_Value As Single) As Integer
'��������� ���������� ������������� ����� � ������� �������
Dim vfi_Temp As Integer

vfi_Temp = Int(afs_Value * (-1)) * (-1)
PF_RoundUp = vfi_Temp

End Function

Public Function CellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber) As Variant
'������� ���������� �������� ������ � ��������� ���������. ���� ����� ������ ���, ���������� 0
    
    On Error GoTo EX
    
    Select Case dataType
        Case Is = visNumber
            CellVal = shp.Cells(cellName).Result(dataType)
        Case Is = visUnitsString
            CellVal = shp.Cells(cellName).ResultStr(dataType)
    End Select
    
    
    
Exit Function
EX:
    CellVal = 0
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
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub


