Attribute VB_Name = "m_tools"
Option Explicit



Public Function GetReadyString(ByVal val As Variant, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'���������� �������������� ������ � ���� prefix & val & postfix. � ������ ���� val=ignore, ���������� ifEmpty
'� ������ ������������� �������������� val � ������ ���������� ifEmpty
    
    On Error GoTo ex
    
    If val = ignore Then
        GetReadyString = ifEmpty
    Else
        GetReadyString = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyString = ifEmpty
End Function
Public Function GetReadyStringA(ByVal elemID As String, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'���������� �������������� ������ � ���� prefix & val & postfix. � ������ ���� val=ignore, ���������� ifEmpty
'� ������ ������������� �������������� val � ������ ���������� ifEmpty
'�������������� ����������� � A, ��� elemID - ��� ������ ���������� ������������ �������
Dim val As Variant
    
    On Error GoTo ex
    
    val = A.Result(elemID)
    If val = ignore Then
        GetReadyStringA = ifEmpty
    Else
        GetReadyStringA = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyStringA = ifEmpty
End Function

Public Sub fixAllGFSShapesC()
'������ ���������� C �� ������� �
Dim shp As Visio.Shape
    
    For Each shp In A.gfsShapes
        SetCellVal shp, "Prop.Unit", Replace(cellVal(shp, "Prop.Unit", visUnitsString), "C", "�")
    Next shp
End Sub

Public Function ClearString(ByVal txt As String) As String
'������� �������� �� ���� ������ � ���������� ����������� �� ����� ����� (���� ������ ������������ �����) ��� �������� ������
Dim tmpVal As Variant
    On Error Resume Next
    txt = Round(txt, 2)
    ClearString = txt
End Function


'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.Version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.Description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub

