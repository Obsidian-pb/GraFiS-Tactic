Attribute VB_Name = "Tools"
Option Explicit


Public Function ListImport(TableName As String, FieldName As String) As String
'������� ��������� ������������ ������ �� ���� ������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object

    On Error GoTo EX

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        
    '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
    With rst
        .MoveFirst
        Do Until .EOF
            List = List & Replace(RSField, Chr(34), "") & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
    ListImport = List

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport"
    ListImport = Chr(34) & " " & Chr(34)
    Set dbs = Nothing
    Set rst = Nothing
End Function

Public Function ListImport2(TableName As String, FieldName As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object, RSField2 As Object

    On Error GoTo EX

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
            "And " & Criteria & _
        "GROUP BY [" & FieldName & "]; "
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        
    '---��������� ���������� ������� � ������ � ���� �� 0 ���������� 0
        If rst.RecordCount > 0 Then
        '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
            With rst
                .MoveFirst
                Do Until .EOF
                    List = List & Replace(RSField, Chr(34), "") & ";"
                    .MoveNext
                Loop
            End With
        Else
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
    ListImport2 = List

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport2"
    ListImport2 = Chr(34) & " " & Chr(34)
    Set dbs = Nothing
    Set rst = Nothing
End Function

Public Function ListImportNum(TableName As String, FieldName As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������ (��� �������� ��������)
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object
    
    On Error GoTo EX
    
'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=0) " & _
            "And " & Criteria & _
        " GROUP BY [" & FieldName & "]; "
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        
    '---��������� ���������� ������� � ������ � ���� �� 0 ���������� 0
        If rst.RecordCount > 0 Then
        '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
            With rst
                .MoveFirst
                Do Until .EOF
                    List = List & Replace(RSField, Chr(34), "") & ";"
                    .MoveNext
                Loop
            End With
        Else
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImportNum = List

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImportNum"
    ListImportNum = Chr(34) & " " & Chr(34)
    Set dbs = Nothing
    Set rst = Nothing
End Function

Public Function ValueImportStr(TableName As String, FieldName As String, Criteria As String) As String
'��������� ��������� �������� ������������� ���� ������� ���������������� ������ ����� ���� �� �������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As String

    On Error GoTo EX

'---���������� ������ � �������������� ����������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            " And " & Criteria & "; "
            
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        
    '---� ������������ � ���������� ������� ���������� �������� �������� ����
        With rst
            .MoveFirst
            ValOfSerch = RSField
        End With
        ValOfSerch = Chr(34) & ValOfSerch & Chr(34)

ValueImportStr = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ValueImportStr"
    ValueImportStr = Chr(34) & " " & Chr(34)
    Set dbs = Nothing
    Set rst = Nothing
End Function


Public Function ValueImportSng(TableName As String, FieldName As String, Criteria As String) As Single
'��������� ��������� �������� ������������� ���� ������� ���������������� ������ ����� ���� �� �������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As Single

    On Error GoTo EX
    
'---���������� ������ � �������������� ����������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]= 0) " & _
            " And " & Criteria & "; "
            
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        
    '---� ������������ � ���������� ������� ���������� �������� �������� ����
        If rst.RecordCount > 0 Then
            With rst
                .MoveFirst
                ValOfSerch = RSField
            End With
        Else
            ValueImportSng = 0
            Set dbs = Nothing
            Set rst = Nothing
            Exit Function
        End If

ValueImportSng = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ValueImportSng"
    ValueImportSng = "0"

    Set dbs = Nothing
    Set rst = Nothing
End Function



Public Function CD_MasterExists(MasterName As String) As Boolean
'������� �������� ������� ������� � �������� ���������
Dim i As Integer

    For i = 1 To Application.ActiveDocument.Masters.Count
        If Application.ActiveDocument.Masters(i).Name = MasterName Then
            CD_MasterExists = True
            Exit Function
        End If
    Next i
    
    CD_MasterExists = False

End Function

Public Sub MasterImportSub(DocName As String, MasterName As String)
'��������� ������� ������� � ������������ � ������
Dim mstr As Visio.Master

    If Not CD_MasterExists(MasterName) Then
        Set mstr = Application.Documents(DocName).Masters(MasterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

End Sub


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

'--------------------------------�������� ����������� ��������� �����---------------------
Public Function IsSelectedOneShape(ShowMessage As Boolean) As Boolean
'---��������� ������ �� ����� ���� ��������� ������
    If Not Application.ActiveWindow.Selection.Count = 1 Then
        If ShowMessage Then MsgBox "�� ������� �� ���� ������ ��� ������� ������ ����� ������", vbInformation
        IsSelectedOneShape = False
        Exit Function
    End If
IsSelectedOneShape = True
End Function
Public Function IsHavingUserSection(ShowMessage As Boolean) As Boolean
'---���������, �� �������� �� ��������� ������ ��� ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        If ShowMessage Then MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ���� �������", vbInformation
        IsHavingUserSection = False
        Exit Function
    End If
IsHavingUserSection = True
End Function
Public Function IsSquare(ShowMessage As Boolean) As Boolean
'---��������� �������� �� ��������� ������ ��������
    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
        If ShowMessage Then MsgBox "��������� ������ �� ����� �������!", vbInformation
        IsSquare = False
        Exit Function
    End If
IsSquare = True
End Function

Public Sub ShowCommonData(shp As Visio.Shape)
'���������� ����� �������� � �������������
Dim common As String

    common = shp.Cells("Prop.Common").ResultStr(0)
    
    f_INPPV_CommonData.ShowData common
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


