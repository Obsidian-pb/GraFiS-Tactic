Attribute VB_Name = "m_Labels"
Public Sub ConnectedShapesLostCheck(ShpObj As Visio.Shape)
'��������� ���������, �� ���� �� ������� ������ � ������� ������������ �������, � ���� ����, �� ������� ��� ���������
On Error GoTo ex
    
    If Not InStr(1, ShpObj.Cells("PinX").FormulaU, "GUARD") > 0 Then
        ShpObj.Delete
    End If
Exit Sub
ex:
    '������
End Sub
