Attribute VB_Name = "m_Connect"
Public Sub ColConn(ShpObj As Visio.Shape)
'��������� �������� ����������� ������ !�������! �� ������ � ������� ��� ��������� (���� ��)
Dim ToShape As Integer

'---������������� ��������� ��������� �� ������
On Error GoTo SubExit

'---���� ������� �� � ���� �� ���������, ��������� �������������
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("Prop.FlowIn").FormulaU = "Sheet." & ToShape & "!Prop.Production" & ""
        ShpObj.Cells("User.WSShapeID").Formula = ToShape
    Else
        ShpObj.Cells("Prop.FlowIn").FormulaU = 0
        ShpObj.Cells("User.WSShapeID").Formula = 0
    End If

Exit Sub

SubExit:
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "ColConn"
End Sub

