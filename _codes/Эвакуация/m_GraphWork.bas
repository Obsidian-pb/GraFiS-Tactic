Attribute VB_Name = "m_GraphWork"
Option Explicit


'------------������ ��� �������� �������� ������ � ������---------------
Public Sub RenumNodes()
'����������������� ����� ����� �����
Dim shp As Visio.Shape
Dim i As Integer
    
    i = 1
    For Each shp In Application.ActivePage.Shapes
        If IsGFSShapeWithIP(shp, indexPers.ipEvacNode) Then
            SetCellVal shp, "Prop.NodeNumber", i
            i = i + 1
        End If
    Next shp
End Sub


Public Sub CalcTimes()
Dim graph As c_WayGraph
Dim controller As c_ControllerGraph
    
    
    If Application.ActiveWindow.Selection.count <> 1 Then
        MsgBox "�� ���� ������ ���� ����� �� �������"
        Exit Sub
    End If
    
    
    Set graph = New c_WayGraph
    Set controller = New c_ControllerGraph
    
    '������ ����
    graph.BuildGraph Application.ActiveWindow.Selection(1)
    
    '������� ������� ����� �����
    controller.SetGraph(graph).ClearGraph.ShapesRefresh
    controller.SetF 0.1       '��������� ������� �������� � ������ "������;�������-�������;������"/"0.1;0.113;0.125"
    controller.SetF InputBox("������� ��������?", "������� ������� �������� ��������", 0.1)
    controller.ResolveGraph_PeopleFlow
    controller.calculate
    controller.ResolveGraph_TimesFlow
    controller.ShapesRefresh
    
    Debug.Print "����� ����� ��������� ��� ����� ������ ���� �����: " & controller.TotalTime
    Debug.Print "����� ��������� �� ���������� ����: " & controller.graph.exitNodes(1).t_flow
'    MsgBox "����� ����� ��������� ��� ����� ������ ���� �����: " & controller.TotalTime & vbNewLine & _
'            "����� ��������� �� ���������� ����: " & controller.graph.exitNodes(1).t_flow
    MsgBox "����� ���������: " & Round(controller.graph.exitNodes(1).t_flow, 1) & " ���."
    
    Set graph = Nothing
End Sub


Public Sub SelectNodes()
'�������� ������ ����� �����
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If IsGFSShapeWithIP(shp, indexPers.ipEvacNode) Then
            Application.ActiveWindow.Select shp, visSelect
        End If
    Next shp
End Sub


Public Sub SeekPlace(ShpObj As Visio.Shape)
'��������� ��������� ���������� ������ ����� � ���������� ��� ������ ���� �����
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim col As Collection

    On Error GoTo ex
'---���������� ���������� �������� ������
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'���������� ��� ������ �� �������� (���� �����)
    For Each OtherShape In Application.ActivePage.Shapes
        If PFB_isDoor(OtherShape) Then
            If OtherShape.HitTest(x, y, 0.01) > 1 Then
                SetCellVal ShpObj, "Prop.WayLen", 0
                SetCellVal ShpObj, "Prop.WayWidth", Round(cellVal(OtherShape, "Width", visMeters), 1)
                SetCellVal ShpObj, "Prop.PeopleHere", 0
                SetCellFrml ShpObj, "Prop.WayClass", "INDEX(1,Prop.WayClass.Format)"
                SetCellFrml ShpObj, "Prop.WayType", "INDEX(1,Prop.WayType.Format)"
                Application.DoCmd 1312
                Exit Sub
            End If
        End If
    Next OtherShape
'���������� ��� ������ �� �������� (���� �����)
    For Each OtherShape In Application.ActivePage.Shapes
        If PFB_isPlace(OtherShape) Then
            If OtherShape.HitTest(x, y, 0.01) > 1 Then
                SetCellVal ShpObj, "Prop.WayLen", Round(cellVal(OtherShape, "Height", visMeters), 0)
'                SetCellVal ShpObj, "Prop.WayWidth", Round(cellVal(OtherShape, "Width", visMeters), 0)
                SetCellVal ShpObj, "Prop.WayWidth", GetWidthByWall(ShpObj)
                SetCellVal ShpObj, "Prop.PeopleHere", cellVal(OtherShape, "Prop.OccupantCount")
                SetCellVal ShpObj, "Prop.PlaceName", cellVal(OtherShape, "Prop.Use", visUnitsString)
                Application.DoCmd 1312
                Exit Sub
            End If
        End If
    Next OtherShape

Application.DoCmd 1312
Set OtherShape = Nothing
Exit Sub
ex:
    Set OtherShape = Nothing
    SaveLog Err, "SeekPlace", ShpObj.Name
End Sub

Public Sub GetShapeLen(ShpObj As Visio.Shape)
    SetCellVal ShpObj, "Prop.EdgeLen", Round(Application.ConvertResult(ShpObj.LengthIU, "in", "m"), 1)
End Sub

Public Function GetNearWallDist(ByRef shp As Visio.Shape) As Single
'������� ���������� �� ��������� �����
Dim wallShp As Visio.Shape
Dim dist As Single
Dim minDist As Single
    
    minDist = 10000
    For Each wallShp In Application.ActivePage.Shapes
        If PFB_isWall(wallShp) Then
            dist = wallShp.DistanceFrom(shp, 0)
            If minDist > dist Then minDist = dist
        End If
    Next wallShp
    
GetNearWallDist = Int(Application.ConvertResult(minDist, "in", "m")) + 1
End Function

Public Function GetWidthByWall(ByRef shp As Visio.Shape) As Single
'���������� ������ ��������� ������ ���������� ���������� �� ��������� ����
    GetWidthByWall = GetNearWallDist(shp) * 2
End Function
