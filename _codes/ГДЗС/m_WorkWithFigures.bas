Attribute VB_Name = "m_WorkWithFigures"
Option Explicit

'--------------------������ ��� ������ � ��������-----------------------------------------


Public Sub PS_GlueToShape(ShpObj As Visio.Shape)
'��������� ����������� �������������� ������ (����� ����) � ������� ������, � ������ ���� ��� �������� _
 ������� �������� ���������
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double
Dim vS_ShapeName As String
Dim shpSize As Double

    On Error GoTo EX

'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---��������� ���OtherShape���� ������ �� ����� ����������� ��������� ������
    '---���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
'    '---���� ������ �������� ������� ���������� � ��� �������� � ��� ������ ����
'        PS_GlueToShape OtherShape
    '---���������, �������� �� ��� ������ ������� �������� ���������
    If GetTypeShape(OtherShape, 104) > 0 Then
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---����������� ������ (����������� �������)
            On Error Resume Next
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." _
                & OtherShape.ID & "!Width*-1.2)*SIN(Sheet." _
                & OtherShape.ID & "!Angle))"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." _
                & OtherShape.ID & "!Width*-1.2)*COS(Sheet." _
                & OtherShape.ID & "!Angle))"
                
                
            ShpObj.Cells("Angle").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Angle+IF(Sheet." _
                & OtherShape.ID & "!User.DownOrient=1,-90 deg,90 deg))"
            ShpObj.Cells("Width").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Width*0.3)"
            ShpObj.Cells("Height").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Height*0.3)"
            ShpObj.Cells("Prop.Unit").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Unit)"
            ShpObj.Cells("Prop.FormingTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.SetTime"
            
            ShpObj.Cells("Prop.Personnel").FormulaU = 1
            ShpObj.Cells("User.ShapeFromID").FormulaU = OtherShape.ID
            ShpObj.Cells("Actions.Release.Invisible").FormulaU = 0
            
            OtherShape.BringToFront
        End If
    End If
    '---���������, �������� �� ��� ������ ������� �������� �������� �����
    If GetTypeShape(OtherShape, 100) > 0 Then
        '---��������� ���������� � ����������� ������
        shpSize = ShpObj.Cells("Height").Result(visInches) / 2
        If OtherShape.HitTest(x, y, shpSize) > 0 Then
            OtherShape.BringToFront
        End If
    End If
Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "PS_GlueToShape"
End Sub

Private Function GetTypeShape(ByRef aO_TergetShape As Visio.Shape, ByVal aI_IndexPers As Integer) As Integer
'������� ���������� ������ ������� ������ � ������ ���� ��� �������� ������� ���������� _
 ��� ������������ �����, � ��������� ������ ������������ 0
Dim vi_TempIndex

GetTypeShape = 0

If aO_TergetShape.CellExists("User.IndexPers", 0) = True And aO_TergetShape.CellExists("User.Version", 0) = True Then
    vi_TempIndex = aO_TergetShape.Cells("User.IndexPers")
    If vi_TempIndex = aI_IndexPers Then
        GetTypeShape = aO_TergetShape.Cells("User.IndexPers")
    End If
End If

End Function


Public Sub PS_ReleaseShape(ShpObj As Visio.Shape, ShapeID As Long)
'��������� ������� ����������� ��������� ������ �� �����������
Dim OtherShape As Visio.Shape

If ShapeID = 0 Then Exit Sub

Set OtherShape = Application.ActivePage.Shapes.ItemFromID(ShapeID)

On Error Resume Next

    ShpObj.Cells("PinX").FormulaForce = ShpObj.Cells("PinX").Result(visNumber)
    ShpObj.Cells("PinY").FormulaForce = ShpObj.Cells("PinY").Result(visNumber)
    ShpObj.Cells("Angle").FormulaForce = ShpObj.Cells("Angle").Result(visNumber)
    ShpObj.Cells("Width").FormulaForce = ShpObj.Cells("Width").Result(visNumber)
    ShpObj.Cells("Height").FormulaForce = ShpObj.Cells("Height").Result(visNumber)
    ShpObj.Cells("Prop.Unit").FormulaForceU = """" & ShpObj.Cells("Prop.Unit").ResultStr(visUnitsString) & """"
    ShpObj.Cells("Prop.SetTime").FormulaForce = ShpObj.Cells("Prop.SetTime").Result(visDate)
    ShpObj.Cells("User.DownOrient").FormulaU = "IF(Angle>-1 deg,1,0)"
    ShpObj.Cells("User.ShapeFromID").FormulaU = 0
    
    ShpObj.Cells("Actions.Release.Invisible").FormulaU = 1
    
    ShpObj.BringToFront

Set OtherShape = Nothing
End Sub

Public Sub MoveMeFront(ShpObj As Visio.Shape)
'����� ���������� ������ ������
    ShpObj.BringToFront
    
'---���������, �� ����������� �� ������ ����� ������ �������� �����
'    PS_GlueToShape ShpObj
End Sub
