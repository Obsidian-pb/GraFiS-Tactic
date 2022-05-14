Attribute VB_Name = "m_MatrixImportExport"
Option Explicit


'---------------------------������ ��� �������� �������/��������----------------------------------
Public Sub SaveMatrixTo(Optional path As String = "")
'�������� ������� �������� ����������� � ������� ������� numpy � csv ����
Dim lay As Variant
Dim s As String
Dim x As Integer
Dim y As Integer

    If path = "" Then
        path = Replace(Application.ActiveDocument.fullName, ".vsdx", ".csv")
        path = Replace(path, ".vsd", ".csv")
    End If
    
'---�������� ������� �������� �����������
    lay = fireModeller.GetOpenSpaceLayer

'---��������� ���� ������� � csv (���� ��� ��� - �������)
    Open path For Output As #1
    
'---��������� ������ �������
    For y = 0 To UBound(lay, 2)
        s = ""
        For x = 0 To UBound(lay, 1)
            s = s & CStr(lay(x, y)) & ","
        Next x
    '---���������� � ����� ����� ���� �������� � ������
        Print #1, Left(s, Len(s) - 1)
    Next y

'---��������� ���� ����
    Close #1
End Sub



