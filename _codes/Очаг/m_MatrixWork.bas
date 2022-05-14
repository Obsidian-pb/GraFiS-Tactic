Attribute VB_Name = "m_MatrixWork"
Option Explicit
'--------------------------------------������ �������/�������� ������----------------------





Public Sub SaveMatrixTo(Optional path As String = "")
'��������� ������� � csv ����
Dim mLayer As Variant
Dim layerString As String
    
    '---�������� ������ ������� ���� �������� �����������
    arrShape fireModeller.GetOpenSpaceLayer
    
    '---�������� ������ ���� �������� ����������� ������� � ����������� ��� � ������
    layerString = Array2DToString(fireModeller.GetOpenSpaceLayer)
    
    '---���� ���� � ����� �� ��� �������, ��������� ���� ��-��������� ��� ���������
    If path = "" Then
        path = Replace(Application.ActiveDocument.fullName, ".vsdx", ".csv")
        path = Replace(path, ".vsd", ".csv")
    End If
    
    '---��������� � ����
    SaveTextToFile layerString, path
    
    '---�������� � ����� ������ �������
     Debug.Print arrShape(fireModeller.GetOpenSpaceLayer)
End Sub

Private Function Array2DToString(arr As Variant) As String
Dim i As Integer
Dim j As Integer
Dim s As String
    

    For j = 0 To UBound(arr, 2)
        For i = 0 To UBound(arr, 1)
            s = s + CStr(arr(i, j)) & ","
        Next i
    Next j

    
Array2DToString = Left(s, Len(s) - 1)
End Function


Public Function arrShape(arr As Variant) As String
'�������� ����������� �������
Dim i As Integer
Dim s As String
    
    On Error Resume Next
    
    For i = 1 To 10
        s = s & UBound(arr, i) + 1 & ","
    Next i

arrShape = Left(s, Len(s) - 1)
End Function



'�����
Public Sub AAA()
Dim a(1, 2) As Integer
    
    a(0, 0) = 10
    a(0, 1) = 20
    a(0, 2) = 30
    a(1, 0) = 40
    a(1, 1) = 50
    a(1, 2) = 60
    
    Debug.Print arrShape(a)
    Debug.Print Array2DToString(a)
End Sub

