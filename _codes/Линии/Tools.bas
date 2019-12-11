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
                        If .Fields(k).Value >= 0 Then
                            '---������������ ������� �������� ������ �������� � ����������� � �� �������� � ��
                            fieldType = .Fields(k).Type
                            If fieldType = 202 Then _
                                ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  '�����
                            If fieldType = 2 Or fieldType = 3 Or fieldType = 4 Or fieldType = 5 Then _
                                ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   '�����
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

'----------------------------------------------------------------------------------------------
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

Public Function ListImport2(TableName As String, FieldName As String, Criteria As String) As String
'������� ��������� ���������� ������ �� ���� ������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object, RSField2 As Object

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
                    List = List & RSField & ";"
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
End Function

Public Function ValueImportStr(TableName As String, FieldName As String, Criteria As String) As String
'��������� ��������� �������� ������������� ���� ������� ���������������� ������ ����� ���� �� �������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As String

'---���������� ������ � �������������� ����������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
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
End Function

Public Function ValueImportSng(TableName As String, FieldName As String, Criteria As String) As Single
'��������� ��������� �������� ������������� ���� ������� ���������������� ������ ����� ���� �� �������
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As Single

'---���������� ������ � �������������� ����������
    '---���������� ������ SQL ��� ������ ������� �� ���� ������
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
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


Public Sub SetFormulaForAll(ShpObj As Visio.Shape, ByVal aS_CellName As String, ByVal aS_NewFormula As String)
'��������� ������������� ���������� ��� ������������ ������
Dim Shp As Visio.Shape

    '���������� ��� ������ � ��������� � ���� ��������� ������ ����� ����� �� ������ - ����������� �� ����� ��������
    For Each Shp In Application.ActiveWindow.Selection
        If Shp.CellExists(aS_CellName, 0) = True Then
            Shp.Cells(aS_CellName).FormulaU = """" & aS_NewFormula & """"
        End If
    Next Shp
End Sub

Public Sub MoveMeFront(ShpObj As Visio.Shape)
'����� ���������� ������ ������
    ShpObj.BringToFront
End Sub

'-----------------------------------------��������� ������ � ��������----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'��������� ������������� ����� �������� ��� ���� ��������� ����� ������ ����
Dim Shp As Visio.Shape
    
    '���������� ��� ������ � ��������� � ���� ��������� ������ ����� ����� �� ������ - ����������� �� ����� ��������
    For Each Shp In Application.ActiveWindow.Selection
        If Shp.CellExists(aS_CellName, 0) = True Then
            Shp.Cells(aS_CellName).Formula = aB_Value
        End If
    Next Shp
    
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
        IsHavingUserSection = True
        Exit Function
    End If
IsHavingUserSection = False
End Function
Public Function IsSquare(ShowMessage As Boolean) As Boolean
'---��������� �������� �� ��������� ������ ��������
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        If ShowMessage Then MsgBox "��������� ������ �� �������� ������� �����!", vbInformation
        IsSquare = True
        Exit Function
    End If
IsSquare = False
End Function
Public Function ClickAndOnSameButton(ClickedButtonName As String) As Boolean
'---��������� �� ������ �� �� �� ������, ��� ������� � ������ ������
'---������ ���� ������ � ������� ���� � �� �� ������
'---����, ���� ��� ������ ������ ��� �� ������ �� ����� ������
Dim v_Cntrl As CommandBarControl
    
'---�������� ��������� ������
    On Error GoTo Tail

'---��������� ������ �� ��������� ������ ��� ������
    For Each v_Cntrl In Application.CommandBars("�����������").Controls
        If v_Cntrl.State = msoButtonDown And v_Cntrl.Caption = ClickedButtonName Then
            ClickAndOnSameButton = True
            Exit Function
        End If
    Next v_Cntrl
    
    ClickAndOnSameButton = False
Exit Function
Tail:
    ClickAndOnSameButton = False
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "NoButtonsON"
End Function

'--------------------------------������----------------------------------------------------------------
Public Function Index(ByVal str As String, ByVal stringArray As String, ByVal delimiter) As Integer
'��� ��� � ��� �������
Dim arr() As String
Dim i As Integer
    
    arr = Split(stringArray, delimiter)
    
    i = 0
    For i = 0 To UBound(arr)
        If arr(i) = str Then
            Index = i
            Exit Function
        End If
    Next i
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
    errString = Now & d & Environ("OS") & d & Environ("HOMEPATH") & d & Environ("APPDATA") & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---���������� � ����� ����� ���� �������� � ������
    Print #1, errString
    
'---��������� ���� ����
    Close #1

End Sub


