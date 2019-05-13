Attribute VB_Name = "Connects"

Public Sub Conn(ShpObj As Visio.Shape)
'��������� �������� ������������� �������� � ������� � ������ � ������ ��� ���������
Dim ToShape As Integer

'---������������� ��������� ��������� �� ������
On Error Resume Next

'---���� ������� �� � ���� �� ���������, ��������� �������������
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.HoseDiameter").FormulaU = "Sheet." & ToShape & "!Prop.HoseDiameter" & ""
        ShpObj.Cells("User.HoseNumber").FormulaU = "Sheet." & ToShape & "!User.HosesNeed" & ""
        ShpObj.Cells("User.WaterExpence").FormulaU = "Sheet." & ToShape & "!Prop.Flow.Value" & ""
        ShpObj.Cells("User.Resistance").FormulaU = "Sheet." & ToShape & "!Prop.HoseResistance.Value" & ""
        ShpObj.Cells("User.LineLenight").FormulaU = "Sheet." & ToShape & "!User.TotalLenight" & ""
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.HoseDiameter").FormulaU = 0
        ShpObj.Cells("User.HoseNumber").FormulaU = 0
        ShpObj.Cells("User.WaterExpence").FormulaU = 0
        ShpObj.Cells("User.Resistance").FormulaU = 0
        ShpObj.Cells("User.LineLenight").FormulaU = 0
    End If

End Sub
