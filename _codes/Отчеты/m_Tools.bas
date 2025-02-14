Attribute VB_Name = "m_Tools"
'-----------------------------------------������ ���������������� �������----------------------------------------------
Option Explicit



Public Function PF_RoundUp(afs_Value As Single) As Integer
'��������� ���������� ������������� ����� � ������� �������
Dim vfi_Temp As Integer

'vfi_Temp = Int(afs_Value * (-1)) * (-1)
'PF_RoundUp = vfi_Temp
    vfi_Temp = Int(afs_Value)
    If vfi_Temp = afs_Value Then
        PF_RoundUp = vfi_Temp
    ElseIf vfi_Temp < afs_Value Then
        PF_RoundUp = vfi_Temp + 1
    End If

End Function

'Public Function cellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, Optional defaultValue As Variant = 0) As Variant
''������� ���������� �������� ������ � ��������� ���������. ���� ����� ������ ���, ���������� 0
'
'    On Error GoTo EX
'
'    If shp.CellExists(cellName, 0) Then
'        Select Case dataType
'            Case Is = visNumber
'                cellVal = shp.Cells(cellName).Result(dataType)
'            Case Is = visUnitsString
'                cellVal = shp.Cells(cellName).ResultStr(dataType)
'            Case Is = visDate
'                cellVal = shp.Cells(cellName).Result(dataType)
'        End Select
'    Else
'        cellVal = defaultValue
'    End If
'
'
'Exit Function
'EX:
'    cellVal = defaultValue
'End Function

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

Public Function SM(ByVal str As String, ByVal count As Integer) As String
'��������������� ������ ��������� ���������� ���
'��������, SM("*",5) => "*****"
Dim i As Integer
Dim strTmp As String
    For i = 1 To count
        strTmp = strTmp & str
    Next i
SM = strTmp
End Function

Public Function SelectedShape() As Visio.Shape
    On Error GoTo ex
    Set SelectedShape = Application.ActiveWindow.Selection(1)
Exit Function
ex:
    Set SelectedShape = Nothing
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


'-------------------------------���������� �������� �����----------------------------------------------
Public Function Sort(ByVal strIn As String, Optional ByVal delimiter As String = ";") As String
'������� ���������� ��������������� ������ � �������� ��������� �������� ����������� delimiter
Dim arrIn() As String
Dim ResultStr As String
Dim arrSize As Integer
Dim gapString As String
Dim i As Integer
Dim j As Integer

    arrIn = Split(strIn, delimiter)
    arrSize = UBound(arrIn)
    
    For i = 0 To arrSize
        For j = i + 1 To arrSize
            If arrIn(j) < arrIn(i) Then
                gapString = arrIn(i)
                arrIn(i) = arrIn(j)
                arrIn(j) = gapString
            End If
        Next j
    Next i
    
    For i = 0 To arrSize
        ResultStr = ResultStr & arrIn(i) & delimiter
    Next i
    
Sort = Left(ResultStr, Len(ResultStr) - Len(delimiter))
End Function

