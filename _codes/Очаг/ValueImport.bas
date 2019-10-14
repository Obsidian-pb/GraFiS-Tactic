Attribute VB_Name = "ValueImport"
'------------------------������ ��� �������� ������� �������� �����-------------------
'------------------------���� �������� �����------------------------------------------
Public Sub GetFactorsByDescription(ShpIndex As Long)
'��������� ������� ������ � �������������� � �������� �������� �� ���� ������ Signs �� �������� ������
Dim dbsE As Database, rsType As Recordset
Dim pth As String
Dim Critria As String, Categorie As String, description As String, IntenseW As Single, speed As Single
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
    Set dbsE = GetDBEngine.OpenDatabase(pth)
    Set rsType = dbsE.OpenRecordset("�_�������������", dbOpenDynaset) '�������� ������ �������

'---���� ����������� ������ � ������ ������ � �� ��� ���������� ������������� ��� �������� ����������
    With rsType
        .FindFirst Criteria
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
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "Sm_ShapeFormShow"
End Sub

