Attribute VB_Name = "Exchange"
'------------------------������ ��� �������� ������� ���-------------------
'-----------------------------------����������� ����-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long)
'����������� ��������� ������� ��� ���������
'---��������� ����������
Dim shp As Visio.Shape
Dim IndexPers As Integer

On Error GoTo ex

'---��������� � ����� ������ ������ ��������� ������ ������
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---��������� ��������� ��������� �������� ��� ��� ������� ������
    Select Case IndexPers
        Case Is = 46
            GetValuesOfCellsFromTable ShpIndex, "����"
        Case Is = 90
            GetValuesOfCellsFromTable ShpIndex, "����"
        Case Is = 49 '��������
            FogRMKGetValuesOfCellsFromTable ShpIndex, "��������"
    
    End Select

Exit Sub
ex:
    SaveLog Err, "GetTTH", CStr(ShpIndex)
End Sub



