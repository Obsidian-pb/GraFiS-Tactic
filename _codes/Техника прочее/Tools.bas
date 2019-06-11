Attribute VB_Name = "Tools"

Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ����� ������ c "�������" �� ���� ������ Signs
Dim dbs As DAO.Database, rsPA As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, PAModel As String, PASet As String
Dim i, k As Integer '������� ��������

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[������] = '" & PAModel & "' And [�����] = '" & PASet & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsPA = dbs.OpenRecordset(TableName, dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rsPA
        .FindFirst Criteria
'MsgBox .RecordCount
    '---���������� ��� ������ ������
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---���������� ��� ���� ������ �������
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---������������ ������� �������� ������ �������� � ����������� � �� �������� � ��
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  '�����
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   '�����
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = False
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = True
                    End If
                    
                End If
                
            Next k
        Next i

    End With

'---��������� ���������� � ��
Set rsPA = Nothing
Set dbs = Nothing

Exit Sub
    '---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
'    MsgBox Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetValuesOfCellsFromTable"
    Set rsPA = Nothing
    Set dbs = Nothing
End Sub


Public Sub GetValuesOfCellsFromTableSea(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ����� ������ c "�������" �� ���� ������ Signs
Dim dbs As DAO.Database, rsPA As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, PAModel As String, PASet As String
Dim i, k As Integer '������� ��������

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[������] = '" & PAModel & "' And [�����] = '" & PASet & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsPA = dbs.OpenRecordset(TableName, dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rsPA
        .FindFirst Criteria
'MsgBox .RecordCount
    '---���������� ��� ������ ������
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---���������� ��� ���� ������ �������
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---������������ ������� �������� ������ �������� � ����������� � �� �������� � ��
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  '�����
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   '�����
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = False
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = True
                    End If
                    
                End If
                
            Next k
        Next i

    End With

'---��������� ���������� � ��
Set rsPA = Nothing
Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
'    MsgBox Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetValuesOfCellsFromTableSea"
    Set rsPA = Nothing
    Set dbs = Nothing

End Sub


Public Sub GetValuesOfCellsFromTableTrain(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ����� ������ c "�������" �� ���� ������ Signs
Dim dbs As DAO.Database, rsPA As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, PAModel As String, PASet As String
Dim i, k As Integer '������� ��������

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[���������] = '" & PAModel & "' And [�����] = '" & PASet & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsPA = dbs.OpenRecordset(TableName, dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rsPA
        .FindFirst Criteria
'MsgBox .RecordCount
    '---���������� ��� ������ ������
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---���������� ��� ���� ������ �������
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---������������ ������� �������� ������ �������� � ����������� � �� �������� � ��
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  '�����
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   '�����
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = False
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = True
                    End If
                    
                End If
                
            Next k
        Next i

    End With

'---��������� ���������� � ��
Set rsPA = Nothing
Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
'    MsgBox Err.Description
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetValuesOfCellsFromTableTrain"
    Set rsPA = Nothing
    Set dbs = Nothing

End Sub

Public Function ListImport(TableName As String, FieldName As String) As String
'������� ��������� ������������ ������ �� ���� ������
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field

    On Error GoTo EX

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
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

Exit Function
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport"
    ListImport = Chr(34) & " " & Chr(34)
End Function

Public Function ListImport2(TableName As String, FieldName As String, FieldName2 As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

    On Error GoTo EX

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "], [" & FieldName2 & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "], [" & FieldName2 & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ') " & _
            "AND (([" & FieldName2 & "])= '" & Criteria & "');"
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
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
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport2 = List

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport2"
    ListImport2 = Chr(34) & " " & Chr(34)
End Function

Public Function ListImport3(TableName As String, FieldName As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

    On Error GoTo EX

'---���������� ����� �������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            "And " & Criteria & _
        "GROUP BY [" & FieldName & "]; "
        
    '---������� ����� ������� ��� ��������� ������
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  '�������� ������ �������
        Set RSField = rst.Fields(FieldName)
        
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
ListImport3 = List

Set dbs = Nothing
Set rst = Nothing

Exit Function
EX:
Set dbs = Nothing
Set rst = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport3"
    ListImport3 = Chr(34) & " " & Chr(34)
End Function

Public Function GetDBEngine() As Object
'Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
Dim engine As Object
    On Error GoTo EX
    Set GetDBEngine = DBEngine
Exit Function
EX:
    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
End Function



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

'--------------------------------���������� ���� ������-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'����� ���������� ���� ���������
Dim errString As String
Const d = " | "

'---��������� ���� ���� (���� ��� ��� - �������)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---��������� ������ ������ �� ������ (���� | �� | Path | APPDATA
    errString = Now & d & Environ("OS") & d & Environ("HOMEPATH") & d & Environ("APPDATA") & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub
