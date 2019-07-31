Attribute VB_Name = "Tools"

Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ����� ������ c "�������" �� ���� ������ Signs
Dim dbs As DAO.Database, rsAD As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, AirDeviceModel As String
Dim i, k As Integer '������� ��������

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
    On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    AirDeviceModel = ShpObj.Cells("Prop.AirDevice").ResultStr(visUnitsString)
'    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.AirDevice.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[������] = '" & AirDeviceModel & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
'    Set dbs = DBEngine.OpenDatabase(pth)
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rsAD
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
    Set rsAD = Nothing
    Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
    MsgBox Err.description
    Set rsAD = Nothing
    Set dbs = Nothing
    SaveLog Err, "GetValuesOfCellsFromTable", "Tablename: " & TableName
End Sub


Public Sub FogRMKGetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'��������� ������� ������ � ��� ���������
Dim dbs As DAO.Database, rsAD As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, FogRMKModel As String
Dim i, k As Integer '������� ��������

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
    On Error GoTo Tail

'---���������� ������ ������������ ������� ����������� ��������
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    FogRMKModel = ShpObj.Cells("Prop.FogRMK").ResultStr(visUnitsString)
'    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.FogRMK.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "������ � ������ " & PASet & " �����������!", vbInformation
            Exit Sub '���� � ������ "������" ������ �������� - ��������� ������������
        End If
    Criteria = "[������] = '" & FogRMKModel & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
'    Set dbs = DBEngine.OpenDatabase(pth)
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ��� �� ��� �������� ����������
    With rsAD
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
Set rsAD = Nothing
Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
    MsgBox Err.description
    Set rsAD = Nothing
    Set dbs = Nothing
    SaveLog Err, "FogRMKGetValuesOfCellsFromTable", "Tablename: " & TableName
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
    '    Set dbs = DBEngine.OpenDatabase(pth)
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
    SaveLog Err, "ListImport", "Tablename: " & TableName
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
    '    Set dbs = DBEngine.OpenDatabase(pth)
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

Exit Function
EX:
    SaveLog Err, "ListImport2", "Tablename: " & TableName
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

Public Sub SetInnerCaptionForAll(ShpObj As Visio.Shape, aS_CellName As String)
'��������� ������������� ���������� ��� ������������ ������
Dim v_Str As String
Dim shp As Visio.Shape

    v_Str = InputBox("������� ����� �������", "��������� �����������")
    '���������� ��� ������ � ��������� � ���� ��������� ������ ����� ����� �� ������ - ����������� �� ����� ��������
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).FormulaU = """" & v_Str & """"
        End If
    Next shp
End Sub

Public Sub MoveMeFront(ShpObj As Visio.Shape)
'����� ���������� ������ ������
    ShpObj.BringToFront
End Sub

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


