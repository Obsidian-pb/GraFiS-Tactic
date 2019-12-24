Attribute VB_Name = "Tools"
'----------------------------������ ������������-----------------------------------


Public Function ListImport(TableName As String, FieldName As String) As String
'������� ��������� ������������ ������ �� ���� ������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object

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
            List = List & RSField & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
    ListImport = List

Set dbs = Nothing
Set rst = Nothing
End Function

Public Function ListImport2(TableName As String, FieldName As String, FieldName2 As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object, RSField2 As Object

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "], [" & FieldName2 & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "], [" & FieldName2 & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ') " & _
            "AND (([" & FieldName2 & "])= '" & Criteria & "');"
        
    '---������� ����� ������� ��� ��������� ������
'        pth = ThisDocument.path & "Signs.fdb"
'        Set dbs = GetDBEngine.OpenDatabase(pth)
'        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
'        Set RSField = rst.Fields(FieldName)
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
        Set RSField = rst.Fields(FieldName)
        Set RSField2 = rst.Fields(FieldName2)
        
    '---��������� ���������� ������� � ������ � ���� �� 0 ���������� 0
        If rst.RecordCount > 0 Then
        '---���� ����������� ������ � ������ ������ � �� ��� ������� ����� �������� ��� ������ ��� �������� ����������
            With rst
                .MoveFirst
                Do Until .EOF
                    List = List & RSField & ";"
                    .MoveNext
                Loop
            End With
        Else
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport2 = List

End Function

'Public Function GetDBEngine() As Object
''Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
'Dim engine As Object
'    On Error GoTo EX
'    Set GetDBEngine = DBEngine
'Exit Function
'EX:
'    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
'End Function



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

Public Sub MasterImportSub(ByVal MasterName As String)
'��������� ������� ������� � ������������ � ������
Dim mstr As Visio.Master

    If Not CD_MasterExists(MasterName) Then
        Set mstr = ThisDocument.Masters(MasterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

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
'---���������, �� �������� �� ��������� ������ ������� �����
    If Application.ActiveWindow.Selection(1).CellExists("User.visObjectType", 0) Then
        If Application.ActiveWindow.Selection(1).Cells("User.visObjectType") = 104 Then
            IsHavingUserSection = True
            Exit Function
        End If
    End If
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



Public Function CheckRushShape() As Boolean
'������� ���������� ������, ���� ������ ����� ������������� � ������ ���������, ���� - ���� ���
    
'---��������� ������ �� ����� ���� ������
    If Application.ActiveWindow.Selection.Count < 1 Then
'        MsgBox "�� ������� �� ���� ������!", vbInformation
        CheckRushShape = False
        Exit Function
    End If
    
'---���������, �� �������� �� ��������� ������ ��� ������� ��� ������ ������� � ������������ ����������
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        MsgBox "��������� ������ ��� ����� ����������� �������� � �� ����� ���� �������� � ���� ���������", vbInformation
        CheckRushShape = False
        Exit Function
    End If
    
'---��������� �������� �� ��������� ������ ��������
    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
        MsgBox "��������� ������ �� ����� ������� � �� ����� ���� �������� � ���� ���������!", vbInformation
        CheckRushShape = False
        Exit Function
    End If
    
CheckRushShape = True
End Function


Public Function IsFirstDrop(ShpObj As Visio.Shape)
'������� ��������� ���������� ������ ������� � ���� �������� ������� ��������� ������� �������� User.InPage
    If ShpObj.CellExists("User.InPage", 0) = 0 Then
        Dim newRowIndex As Integer
        newRowIndex = ShpObj.AddNamedRow(visSectionUser, "InPage", visRowUser)
        ShpObj.CellsSRC(visSectionUser, newRowIndex, 0).Formula = 1
        ShpObj.CellsSRC(visSectionUser, newRowIndex, visUserPrompt).FormulaU = """+"""
        
        IsFirstDrop = True
    Else
        IsFirstDrop = False
    End If
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


