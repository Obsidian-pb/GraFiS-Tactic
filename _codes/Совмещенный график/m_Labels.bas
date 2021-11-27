Attribute VB_Name = "m_Labels"
Option Explicit
'------------------------------������ ��� ������ � �������� �������� � �������-------------------


Public Sub SeekGraphicForLabel(ShpObj As Visio.Shape)
'��������� ��������� �������� � ������� �� � �������� � ���� (��� ������ ���������)
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim Col As Collection
Dim ShapeType As Integer

On Error GoTo EX

'---���������� ���������� �������� ������
x = ShpObj.Cells("EndX").Result(visInches)
y = ShpObj.Cells("Endy").Result(visInches)
'---���������� ��� �������� ������
ShapeType = ShpObj.Cells("User.IndexPers")

'���������� ��� ������ �� ��������
For Each OtherShape In Application.ActivePage.Shapes
    If OtherShape.CellExists("User.IndexPers", 0) = True And OtherShape.CellExists("User.Version", 0) = True Then
        If OtherShape.Cells("User.IndexPers") = 122 And OtherShape.HitTest(x, y, 0.01) > 1 Then  '���� ������ - ����� ������ ��
            '---��������� �������� ������� �� ���������� ���� �������
                ShpObj.Cells("User.FireTime").FormulaU = "Sheet." & OtherShape.ID & "!User.FireTime"
                ShpObj.Cells("User.FireMax").FormulaU = "Sheet." & OtherShape.ID & "!Prop.FireMax"
                ShpObj.Cells("User.TimeMax").FormulaU = "Sheet." & OtherShape.ID & "!Prop.TimeMax"
                ShpObj.Cells("User.ParentGraphWidth").FormulaU = "Sheet." & OtherShape.ID & "!Width"
                ShpObj.Cells("User.ParentGraphHeight").FormulaU = "Sheet." & OtherShape.ID & "!Height"
                ShpObj.Cells("User.WaterIntense").FormulaU = "Sheet." & OtherShape.ID & "!User.WaterIntense"
                ShpObj.Cells("User.FontSizeCaption").FormulaU = "Sheet." & OtherShape.ID & "!User.FontSizeCaption"
                ShpObj.Cells("User.LineWeightLines").FormulaU = "Sheet." & OtherShape.ID & "!User.LineWeightLines"
                
                ShpObj.Cells("User.X0").FormulaU = "Sheet." & OtherShape.ID & "!PinX-Sheet." & _
                        OtherShape.ID & "!LocPinX"
                ShpObj.Cells("User.Y0").FormulaU = "Sheet." & OtherShape.ID & "!PinY-Sheet." & _
                        OtherShape.ID & "!LocPinY"

            '---��� �������������� ����������� ���������� ������� �� ���������
            Set OtherShape = Nothing
            Exit Sub
        End If
    End If
Next OtherShape

'� ������, ���� �� � ����� ������� ���� ������� ���������� �� �������, ����������� ��������
                ShpObj.Cells("User.FireTime").FormulaU = "DateValue(" & ShpObj.Cells("User.FireTime").Result(visNumber) & ")"
                ShpObj.Cells("User.FireMax").FormulaU = ShpObj.Cells("User.FireMax")
                ShpObj.Cells("User.TimeMax").FormulaU = ShpObj.Cells("User.TimeMax").FormulaU
                '!!!��� ����� ���� ������
                ShpObj.Cells("User.ParentGraphWidth").Formula = ShpObj.Cells("User.ParentGraphWidth").Result(visMeters) & "m"
                ShpObj.Cells("User.ParentGraphHeight").Formula = ShpObj.Cells("User.ParentGraphHeight").Result(visMeters) & "m"
                '!!!��� ����� ���� ������
                ShpObj.Cells("User.WaterIntense").Formula = ShpObj.Cells("User.WaterIntense")
                ShpObj.Cells("User.FontSizeCaption").Formula = ShpObj.Cells("User.FontSizeCaption")
                ShpObj.Cells("User.ArrowsSize").Formula = ShpObj.Cells("User.ArrowsSize")
                ShpObj.Cells("User.LineWeightLines").Formula = ShpObj.Cells("User.LineWeightLines")
                
                ShpObj.Cells("User.X0").Formula = ShpObj.Cells("User.X0")
                ShpObj.Cells("User.Y0").Formula = ShpObj.Cells("User.Y0")

                
ShpObj.BringToFront


Set OtherShape = Nothing
EX:
Set OtherShape = Nothing
End Sub


Public Sub SetTime(ByRef shp As Visio.Shape)
    On Error Resume Next
    
    Application.DoCmd 1312
    '����� �� ��������� ����� ���������� � ��������� ��� ��� ������������ ����
    Debug.Print shp.Cells("Prop.MilestoneType").ResultStr(visUnitsString) & ": " & "����� �� ��������� ����� ���������� � ��������� ��� ��� ������������ ����"
End Sub
