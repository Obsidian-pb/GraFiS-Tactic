Attribute VB_Name = "Exchange"
'------------------------������ ��� �������� ������� ���-------------------
'-----------------------------------����������� ����-------------------------------------------------
Public Sub GetTTH(shp As Visio.Shape)
'����������� ��������� ������� ��� �����������
'---��������� ����������
'Dim shp As Visio.Shape
Dim IndexPers As Integer

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
'    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
    Select Case IndexPers
        Case Is = 73 '������ �� ���������� ����
            GetValuesOfCellsFromTable shp, "[�_���������� ������]"
        Case Is = 74 '�����
            GetValuesOfCellsFromTable shp, "[�_���������� ������]"
        Case Is = 30 '�������
            GetValuesOfCellsFromTableSea shp, "�_����"
        Case Is = 31 '�����
            GetValuesOfCellsFromTableSea shp, "�_����"
        Case Is = 24 '������
            GetValuesOfCellsFromTableTrain shp, "�_������"
        Case Is = 28 '���������
            GetValuesOfCellsFromTable shp, "�_���������"
        Case Is = 25 '�������
            GetValuesOfCellsFromTable shp, "�_��������"
        Case Is = 26 '�������-�������
            GetValuesOfCellsFromTable shp, "�_��������"
        Case Is = 27 '��������
            GetValuesOfCellsFromTable shp, "�_���������"
            
            
    End Select

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetTTH"
End Sub



