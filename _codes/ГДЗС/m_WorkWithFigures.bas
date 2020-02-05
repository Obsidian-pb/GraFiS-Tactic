Attribute VB_Name = "m_WorkWithFigures"
Option Explicit



'--------------------������ ��� ������ � ��������-----------------------------------------
Public Sub MoveMeFront(ShpObj As Visio.Shape)
'����� ���������� ������ ������
    ShpObj.BringToFront
    
'---���������, �� ����������� �� ������ ����� ������ �������� �����
    PS_GlueToHose ShpObj

End Sub

Public Sub PS_GlueToShape(ShpObj As Visio.Shape)
'��������� ����������� �������������� ������ (����� ����) � ������� ������, � ������ ���� ��� �������� _
 ������� �������� ���������
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double

    On Error GoTo EX

'---���������� ���������� � ������ �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---��������� ���OtherShape���� ������ �� ����� ����������� ��������� ������
    '---���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
'    '---���� ������ �������� ������� ���������� � ��� �������� � ��� ������ ����
'        PS_GlueToShape ShpObj
    '---���������, �������� �� ��� ������ ������� �������� ���������
        If IsShapeGraFiSType(OtherShape, Array(104)) Then
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
    Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "PS_GlueToShape"
End Sub

Public Sub PS_GlueToHose(ShpObj As Visio.Shape)
'��������� ����������� �������������� ������ (����� ����) � ������� ������, � ������ ���� ��� �������� _
 ������� �������� �����
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double
Dim vS_ShapeName As String

Dim shpSize As Double
Dim curHoseShape As Visio.Shape
Dim curHoseDistance As Double
Dim newHoseDistance As Double

Dim curHoseShapeID As Integer

    On Error GoTo EX

'---��������� ������� �� � ������ ������ ����������� ���� (��� �������� ����� ������������ ����� ����)
    If ShpObj.CellExists("User.ShapeHoseID", 0) = False Then Exit Sub

'---���������� ���������� � ������ �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)
    shpSize = ShpObj.Cells("Height").Result(visInches) / 2
    curHoseDistance = shpSize * 1.01

'---��������� ������� ������ �� ����� �����������
    '---���������� ��� ������ �� ��������
    For Each OtherShape In Application.ActivePage.Shapes
    '    '---���� ������ �������� ������� ���������� � ��� �������� � ��� ������ ����
    '        PS_GlueToHose ShpObj
        '---���������, �������� �� ��� ������ ������� �������� �������� �����
        If IsShapeGraFiSType(OtherShape, Array(100)) Then
            '---���� ��������, ��������� �������� �� ��� � ������� shpSize �� Pin ������ ����� ����
            If OtherShape.HitTest(x, y, shpSize) > 0 Then
                newHoseDistance = OtherShape.DistanceFromPoint(x, y, 0)
                If curHoseDistance > newHoseDistance Then
                    Set curHoseShape = OtherShape
                    curHoseDistance = newHoseDistance
                End If
            End If
        End If
    Next OtherShape

'---���� ���� ������� ������ �������� �����, � ������� ����� ��������� �����
    If Not curHoseShape Is Nothing Then
        curHoseShapeID = ShpObj.Cells("User.ShapeHoseID").Result(visNumber)
    '---1 ��������� ���� �� ����� ��������� � �����
        If curHoseShapeID = 0 Then
        '---���
            SetHoseLineGDZSStatus curHoseShape, ShpObj, True
        Else
        '---��
        '---2 ��������� ���� �� ����� ��������� � ������ �����
            If curHoseShapeID <> curHoseShape.ID Then
            '---������� �������� � ����� ������� ����� � ������ �������� ��� �������
                SetHoseLineGDZSStatus Application.ActivePage.Shapes.ItemFromID(curHoseShapeID), ShpObj, False
                SetHoseLineGDZSStatus curHoseShape, ShpObj, True
            End If
        End If
        ShpObj.Cells("User.ShapeHoseID").Formula = curHoseShape.ID
        
    '---3 ������������ ������ � ������������ � ���������� ����� � ���������� �� �����
        '#TEMP:
        RotateAtHoseLine ShpObj, NewVectorXY(x, y), curHoseShape
        '#END TEMP
    Else
        curHoseShapeID = ShpObj.Cells("User.ShapeHoseID").Result(visNumber)
        '---1 ��������� ���� �� ����� ��� ��������� � �����
        If curHoseShapeID <> 0 Then
            Set curHoseShape = GetShapeByID(curHoseShapeID)
            '---���� ���� ���������, � ����� � ������� ��� ���� ��������� �� �������
            If Not curHoseShape Is Nothing Then
                SetHoseLineGDZSStatus curHoseShape, ShpObj, False
            End If
        End If
        '---���������, ��� ����� ������ �� �������� � ������
        ShpObj.Cells("User.ShapeHoseID").Formula = 0
    End If


Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    MsgBox "� ���� ���������� ��������� ��������� ������! ���� ��� ����� ����������� - ���������� � ������������."
    SaveLog Err, "PS_GlueToHose"
End Sub

Private Sub SetHoseLineGDZSStatus(ByRef shp As Visio.Shape, ByRef shpGDZS As Visio.Shape, ByVal isGDZS As Boolean)
'������������� ������ �������� ���� � ������ ��������������� � ��� (������ ��� �������)
Dim con As Visio.Connect
Dim extShp As Visio.Shape

    If isGDZS Then
        For Each con In shp.Connects
            Set extShp = con.ToSheet
            If IsShapeGraFiSType(extShp, Array(34, 35, 36, 37, 38, 39)) Then
                '��������� ��������, ��� �� ������� �������� ��������� ����� ����
                extShp.Cells("Prop.Personnel").Formula = 0
                extShp.Cells("Prop.Unit").FormulaU = """" & shpGDZS.Cells("Prop.Unit").ResultStr(visUnitsString) & """"
            End If
        Next con
    Else
        For Each con In shp.Connects
            Set extShp = con.ToSheet
            If IsShapeGraFiSType(extShp, Array(34, 35, 36, 37, 38, 39)) Then
                '��������� ��������, ��� �� ������� �� �������� ��������� ����� ����
                extShp.Cells("Prop.Personnel").FormulaU = "IF(STRSAME(Prop.TTHType," & Chr(34) & _
                    "�����������" & Chr(34) & "),IF(Prop.DiameterInS>50,2,1),IF(Prop.DiameterIn>50,2,1))"
            End If
        Next con
    End If

End Sub


Public Sub PS_ReleaseShape(ShpObj As Visio.Shape, ShapeID As Long)
'��������� ������� ����������� ����� ���� �� �������� ����������
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


'-----------------------------------------������� ������ ����� �� �������� �����----------------------------------------------
Private Sub RotateAtHoseLine(ByRef gdzsShp As Visio.Shape, ByRef curPoint As c_Vector, ByRef hoseLineShp As Visio.Shape)
'�������� ��������� �������� ������ �� �������� �����
Dim newCenter As c_Vector
Dim vector1 As c_Vector
Dim vector2 As c_Vector
Dim vector3 As c_Vector
    
    On Error GoTo EX
    
'    Application.EventsEnabled = False
    
    '1 ������� ��������� ����� �� �������� �����
    Set newCenter = FindNearestHoseLinePoint(curPoint, hoseLineShp)
    If newCenter Is Nothing Then Exit Sub
    
    '���������, �� ��������� �� ��� ������ � ������ �����
    If curPoint.IsSame(newCenter, 0.1) Then Exit Sub
    
    '2 ��������� ������ � ��������� �����
    gdzsShp.Cells("PinX").Formula = str(newCenter.x)
    gdzsShp.Cells("PinY").Formula = str(newCenter.y)
    
    '3 ���� ��� ����� �� �����
    Set vector1 = GetPointOnLineShape(curPoint, hoseLineShp, gdzsShp.Cells("Height").Result(visInches) / 2, 0)
    Set vector2 = GetPointOnLineShape(curPoint, hoseLineShp, gdzsShp.Cells("Height").Result(visInches) / 2, vector1.segmentNumber + 1)
    Set vector3 = NewVectorXY(vector1.x - vector2.x, vector1.y - vector2.y)
    gdzsShp.Cells("Angle").Formula = str(vector3.Angle) & "deg"
EX:
'    Application.EventsEnabled = True
End Sub


Public Function FindNearestHoseLinePoint(ByRef curPoint As c_Vector, ByRef hoseLineShp As Visio.Shape) As c_Vector
'������� ��������� ����� �������� �����
Dim distance As Double

    '1 ������� ���������� �� �������� �����
    distance = hoseLineShp.DistanceFromPoint(curPoint.x, curPoint.y, visSpatialIncludeDataGraphics)
    '2 �� ��������� ���������� ������� �����
    Set FindNearestHoseLinePoint = GetPointOnLineShape(curPoint, hoseLineShp, distance)
    
End Function

