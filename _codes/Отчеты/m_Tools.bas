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

Public Function IsGFSShape(ByRef shp As Visio.Shape) As Boolean
'������� ���������� True, ���� ������ �������� ������� ������
Dim i As Integer
    
    '���������, �������� �� ������ ������� ������
    If shp.CellExists("User.IndexPers", 0) = True Then
        IsGFSShape = True
        Exit Function
    End If
    
IsGFSShape = False
End Function

Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPreses As Variant) As Boolean
'������� ���������� True, ���� ������ �������� ������� ������ � ����� ���������� ����� ����� ������ (gfsIndexPreses) ������������ IndexPers ������ ������
'�������������� ��� ���������� ������ ��� ��������� �� ��, ��������� �� ��� ��� � ������� ������. � ������, ���� � ������ ��� ������ User.IndexPers _
'���������� ������ ��������� ������� ������� False
'������ �������������: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
'                 ���: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
Dim i As Integer
Dim indexPers As Integer
    
    On Error GoTo EX
    
    '���������, �������� �� ������ ������� ���������� ����
    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
    Select Case TypeName(gfsIndexPreses)
        Case Is = "Long"    '���� �������� ������������ ��������
            If gfsIndexPreses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Variant()"   '���� ������� ������
            For i = 0 To UBound(gfsIndexPreses)
                If gfsIndexPreses(i) = indexPers Then
                    IsGFSShapeWithIP = True
                    Exit Function
                End If
            Next i
        Case Else
            IsGFSShapeWithIP = False
    End Select

IsGFSShapeWithIP = False
Exit Function
EX:
    IsGFSShapeWithIP = False
    SaveLog Err, "m_Tools.IsGFSShapeWithIP"
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


