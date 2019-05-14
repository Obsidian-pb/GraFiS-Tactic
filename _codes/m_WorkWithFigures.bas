Attribute VB_Name = "m_WorkWithFigures"
'----------------------------� ������ �������� ��������� ������ � ������� ������--------------------------------------
Option Explicit

Public Sub PS_GlueToShape(ShpObj As Visio.Shape)
'��������� ����������� �������������� ������ (�������� �����) � ������� ������, � ������ ���� ��� �������� _
 ������� ��, ���, �� ��� ������
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double
Dim vS_ShapeName As String

    On Error GoTo EX

'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---��������� ���OtherShape���� ������ �� ����� ����������� ��������� ������
    '---���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
'    '---���� ������ �������� ������� ���������� � ��� �������� � ��� ������ ����
'        PS_GlueToShape OtherShape
    '---���������, �������� �� ��� ������ ������� ��, ���, �� ��� ������
    If GetTypeShape(OtherShape) > 0 Then
        '---��������� ���������� � ����������� ������
'        OtherShape.XYFromPage x, y, x, y
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---����������� ������ (����������� �������)
            On Error Resume Next
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." _
                & OtherShape.ID & "!Width*0.55)*SIN(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." _
                & OtherShape.ID & "!Width*0.55)*COS(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("Angle").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Angle+IF(Sheet." _
                & OtherShape.ID & "!User.DownOrient=1,-20 deg,20 deg))"
            ShpObj.Cells("Width").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Width*0.575)"
            ShpObj.Cells("Height").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Height*0.583)"
            ShpObj.Cells("Prop.Unit").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Unit)"
            ShpObj.Cells("Prop.SetTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.ArrivalTime"
            ShpObj.Cells("User.DownOrient").FormulaU = "Sheet." & OtherShape.ID & "!User.DownOrient"
            ShpObj.Cells("User.ShapeFromID").FormulaU = OtherShape.ID
            
            OtherShape.Cells("User.GFS_OutLafet").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.PodOut),0," & _
                "Sheet." & ShpObj.ID & "!User.PodOut)"
            OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.Head),0," & _
                "Sheet." & ShpObj.ID & "!User.Head)"
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



Private Function GetTypeShape(ByRef aO_TergetShape As Visio.Shape) As Integer
'������� ���������� ������ ������� ������ � ������ ���� ��� �������� ������� ���������� _
 ��� ��������� ��������� ������, � ��������� ������ ������������ 0
Dim vi_TempIndex

GetTypeShape = 0

If aO_TergetShape.CellExists("User.IndexPers", 0) = True And aO_TergetShape.CellExists("User.Version", 0) = True Then
    vi_TempIndex = aO_TergetShape.Cells("User.IndexPers")
    If vi_TempIndex = 1 Or vi_TempIndex = 2 Or vi_TempIndex = 9 Or vi_TempIndex = 10 Or _
        vi_TempIndex = 11 Or vi_TempIndex = 20 Then
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
    
    OtherShape.Cells("User.GFS_OutLafet").FormulaU = 0
    OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = 0
    ShpObj.BringToFront

Set OtherShape = Nothing
End Sub
