Attribute VB_Name = "Exchange"
'------------------------������ ��� �������� ������� ���-------------------
'-----------------------------------����������� ����-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long)
'����������� ��������� ������� ��� �����������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer

    On Error GoTo EX

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
    Select Case IndexPers
        Case Is = 73 '������ �� ���������� ����
            GetValuesOfCellsFromTable ShpIndex, "�_���������� ������"
        Case Is = 74 '�����
            GetValuesOfCellsFromTable ShpIndex, "�_���������� ������"
        Case Is = 30 '�������
            GetValuesOfCellsFromTableSea ShpIndex, "�_����"
        Case Is = 31 '�����
            GetValuesOfCellsFromTableSea ShpIndex, "�_����"
        Case Is = 24 '������
            GetValuesOfCellsFromTableTrain ShpIndex, "�_������"
        Case Is = 28 '���������
            GetValuesOfCellsFromTable ShpIndex, "�_���������"
        Case Is = 25 '�������
            GetValuesOfCellsFromTable ShpIndex, "�_��������"
        Case Is = 26 '�������-�������
            GetValuesOfCellsFromTable ShpIndex, "�_��������"
        Case Is = 27 '��������
            GetValuesOfCellsFromTable ShpIndex, "�_���������"
            
            
    End Select

Exit Sub
EX:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "GetTTH"
End Sub



