Attribute VB_Name = "Exchange"
'------------------------������ ��� �������� ������� ���-------------------
'-----------------------------------����������� ����-------------------------------------------------
Public Sub GetTTH(shp As Visio.Shape)
'����������� ��������� ������� ��� �����������
'---��������� ����������
Dim IndexPers As Integer

'---��������� � ����� ������ ������ ��������� ������ ������
    IndexPers = shp.Cells("User.IndexPers")
    
'---��������� ��������� ��������� �������������� ������ ������(�����) ��� ������� ������
Select Case IndexPers
    Case Is = 1
        GetValuesOfCellsFromTable shp, "�_������������"
    Case Is = 2
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 3
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 4
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 5
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 6
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 7
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 8
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 9
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 10
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 11
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 12
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 13
        GetValuesOfCellsFromTable shp, "�_����"
    Case Is = 14
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 15
        GetValuesOfCellsFromTable shp, "�_�����"
    Case Is = 16
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 17
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 18
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 19
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 20
        GetValuesOfCellsFromTable shp, "�_��"
    Case Is = 161
        GetValuesOfCellsFromTable shp, "�_���"
    Case Is = 162
        GetValuesOfCellsFromTable shp, "�_����"
    Case Is = 163
        GetValuesOfCellsFromTable shp, "�_���"
        
        
        
End Select



End Sub


