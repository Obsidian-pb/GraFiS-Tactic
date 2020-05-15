Attribute VB_Name = "ValueImport"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub GetFactorsByDescription(ShpIndex As Long)
'��������� ������� ������ � �������������� � �������� �������� �� ���� ������ Signs �� �������� ������
Dim dbsE As Object, rsType As Object
Dim pth As String
Dim SQLQuery As String, Critria As String, Categorie As String, description As String, IntenseW As Single, speed As Single
Dim shp As Visio.Shape

'---���������� �������� � ������ ������ �������� ������� �������� ���������� ������
    On Error GoTo Tail

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---���������� �������� ������ ������ � ������ ������
    Categorie = shp.Cells("Prop.FireCategorie").ResultStr(visUnitsString)
    description = shp.Cells("Prop.FireDescription").ResultStr(visUnitsString)
    Criteria = "[���������] = '" & Categorie & "' And [��������] = '" & description & "'"
    
'---������� ���������� � �� Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbsE = CreateObject("ADODB.Connection")
    dbsE = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
    dbsE.Open
    Set rsType = CreateObject("ADODB.Recordset")
    SQLQuery = "SELECT * From �_�������������"
    rsType.Open SQLQuery, dbsE, 3, 1

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ������������� ��� �������� ����������
    With rsType
        .Filter = Criteria
        If .RecordCount > 0 Then
            .MoveFirst
            If ![�����������������������] > 0 Then '���� �������� ������������� ������ ���� � �� ���
                Intense = ![�����������������������]
            Else
                MsgBox "��������� �������� ������������� ������ ���� ��� ������� �������� � ���� ������ �����������! " & _
                    "������� �� ��������� ����� ��������� �������� 0�/�*�.��.."
                Intense = 0
            End If
            
            If ![������������] > 0 Then '���� �������� �������� � �� ���
                speed = ![������������]
            Else
                MsgBox "��������� �������� �������� �������� ��������������� ���� ��� ������� �������� � ���� ������ �����������! " & _
                    "������� �� ��������� ����� ��������� �������� 0�/���."
                speed = 0
            End If
        End If
    End With
    
'---����������� ���������� �������� �������
        shp.Cells("Prop.WaterIntense").FormulaU = str(Intense)
        shp.Cells("Prop.FireSpeedLine").FormulaU = str(speed)
    
'---��������� ���������� � ��
rsType.Close
dbsE.Close
Set dbs = Nothing

Exit Sub

'---� ������ ������ �������� ������� �������� ���������� ������, ����������� ���������
Tail:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������.", , ThisDocument.Name
    SaveLog Err, "Sm_ShapeFormShow"
End Sub

