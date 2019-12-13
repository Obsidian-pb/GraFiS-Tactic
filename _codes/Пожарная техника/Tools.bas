Attribute VB_Name = "Tools"
Option Explicit


Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ����� ������ c "�������" �� ���� ������ Signs
Dim dbs As Object, rst As Object
Dim pth As String
Dim ShpObj As Visio.Shape
Dim SQL As String, Criteria As String, PAModel As String, PASet As String
Dim i, k As Integer '������� ��������
Dim fieldType As Integer

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[������] = '" & PAModel & "' And [�����] = '" & PASet & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = CreateObject("ADODB.Connection")
    dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
    dbs.Open
    Set rst = CreateObject("ADODB.Recordset")
    SQL = "SELECT * From " & TableName
    rst.Open SQL, dbs, 3, 1

    

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rst
        .Filter = Criteria
        If .RecordCount > 0 Then
            .MoveFirst
        '---���������� ��� ������ ������
            For i = 0 To ShpObj.RowCount(visSectionProp) - 1
            '---���������� ��� ���� ������ �������
                For k = 0 To .Fields.Count - 1
                    If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                        .Fields(k).Name Then
                        If .Fields(k).value >= 0 Then
                            '---������������ ������� �������� ������ �������� � ����������� � �� �������� � ��
                            fieldType = .Fields(k).Type
                            If fieldType = 202 Then _
                                ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).value & """"  '�����
                            If fieldType = 2 Or fieldType = 3 Or fieldType = 4 Or fieldType = 5 Then _
                                ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).value)   '�����
                        Else
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        End If
                        
                    End If
                    
                Next k
            Next i
        End If
    End With

'---��������� ���������� � ��
Set rst = Nothing
Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
    MsgBox Err.description
    Set rst = Nothing
    Set dbs = Nothing
    SaveLog Err, "GetValuesOfCellsFromTable"

End Sub


Public Function RowIndex(RName As String) As Integer
'������� ���������� ������ ������ � ������ Prop ����-����� ��������� �� ����� RName
Dim RowInd As Integer, RowCount As Integer

For RowInd = 0 To ActiveDocument.DocumentSheet.RowCount(visSectionProp) - 1
    If ActiveDocument.DocumentSheet.CellsSRC(visSectionProp, RowInd, 3).RowNameU = RName Then
        RowIndex = RowInd
        'Exit For
    End If
Next RowInd

End Function

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

    On Error GoTo Tail

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
Exit Function
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ListImport2"
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
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub


'-----------------------------------------������� �������� �������� ������----------------------------------------------
Public Function IsShapeLinkedToDataAndDropFirst(ByRef shp As Visio.Shape) As Boolean
IsShapeLinkedToDataAndDropFirst = False

    If IsShapeLinked(shp) And shp.CellExists("User.InPage", 0) = False Then
        IsShapeLinkedToDataAndDropFirst = True
        Exit Function
    End If
End Function
Public Function IsShapeLinked(ByRef shp As Visio.Shape) As Boolean
Dim rst As Visio.DataRecordset
Dim propIndex As Integer
    
    IsShapeLinked = False
    
    For Each rst In Application.ActiveDocument.DataRecordsets
        For propIndex = 0 To shp.RowCount(visSectionProp)
            If shp.IsCustomPropertyLinked(rst.ID, propIndex) Then
                IsShapeLinked = True
                Exit Function
            End If
        Next propIndex
    Next rst
End Function


