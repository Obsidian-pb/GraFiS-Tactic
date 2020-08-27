Attribute VB_Name = "Tools"
Option Explicit

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

'------------------------------������ � ��������--------------------
Public Function IsGFSShape(ByRef shp As Visio.Shape, Optional ByVal useManeure As Boolean = True) As Boolean
'������� ���������� True, ���� ������ �������� ������� ������
Dim i As Integer
    
'    If shp.CellExists("User.IndexPers", 0) = True and shp.CellExists("User.Version", 0) = True Then        '�������� - ����� �� ������ ���� ������
    '���������, �������� �� ������ ������� ������
    If useManeure Then      '���� ����� ��������� �������� �� ������
        If shp.CellExists("User.IndexPers", 0) = True Then
            '���� ������� ������ ����� ������� � �� �������� ����������, ���
            If shp.CellExists("Actions.MainManeure", 0) = True Then
                If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
                    IsGFSShape = True       '������ ������ � �� �����������
                Else
                    IsGFSShape = False      '������ ������ � �����������
                End If
            Else
                IsGFSShape = True       '������ ������ � �� ����� ������ ������
            End If
        Else
            IsGFSShape = False      '������ �� ������
        End If
    Else                    '���� �� ����� ��������� �������� �� ������
'        If shp.CellExists("User.IndexPers", 0) = True Then
'            IsGFSShape = True       '������ ������
'        Else
'            IsGFSShape = False      '������ �� ������
'        End If
        IsGFSShape = shp.CellExists("User.IndexPers", 0)
    End If

End Function

Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPerses As Variant, Optional needGFSCheck As Boolean = False) As Boolean
'������� ���������� True, ���� ������ �������� ������� ������ � ����� ���������� ����� ����� ������ (gfsIndexPreses) ������������ IndexPers ������ ������
'�� ��������� �������������� ��� ���������� ������ ��� ��������� �� ��, ��������� �� ��� � ������� ������. � ������, ���� � ������ ��� ������ User.IndexPers _
'���������� ������ ��������� ������� ������� False
'������ �������������: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
'                 ���: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
Dim i As Integer
Dim indexPers As Integer
    
    On Error GoTo ex
    
    '���� ���������� ��������������� �������� �� ��������� ������ � ������:
    If needGFSCheck Then
        If Not IsGFSShape(shp) Then
            IsGFSShapeWithIP = False
            Exit Function
        End If
    End If
    
    '���������, �������� �� ������ ������� ���������� ����
    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
    Select Case TypeName(gfsIndexPerses)
        Case Is = "Long"    '���� �������� ������������ �������� Long
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Integer"    '���� �������� ������������ �������� Integer
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Variant()"   '���� ������� ������
            For i = 0 To UBound(gfsIndexPerses)
                If gfsIndexPerses(i) = indexPers Then
                    IsGFSShapeWithIP = True
                    Exit Function
                End If
            Next i
        Case Else
            IsGFSShapeWithIP = False
    End Select

IsGFSShapeWithIP = False
Exit Function
ex:
    IsGFSShapeWithIP = False
    SaveLog Err, "m_Tools.IsGFSShapeWithIP"
End Function

Public Function AddToColl(ByRef coll As Collection, ByVal v As Variant) As Collection
'��������� ������� � ���������
On Error GoTo ex
    
    coll.Add v, str(v)
    
'Exit Function
ex:
    Set AddToColl = coll
End Function

Public Function ColToArr(ByRef coll As Collection) As Variant
'���������� ������ �� ���������
Dim v As Variant
Dim i As Integer
Dim arr As Variant
    
    ReDim arr(coll.Count)
    i = 0
    For Each v In coll
        arr(i) = v
        i = i + 1
    Next v
    
ColToArr = arr
End Function
