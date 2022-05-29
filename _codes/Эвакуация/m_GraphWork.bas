Attribute VB_Name = "m_GraphWork"
Option Explicit


'------------Модуль для хранения процедур работы с графом---------------
Public Sub RenumNodes()
'Перенумеровывание фигур узлов графа
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
        MsgBox "Ни одна фигура узла графа не выбрана"
        Exit Sub
    End If
    
    
    Set graph = New c_WayGraph
    Set controller = New c_ControllerGraph
    
    'Строим граф
    graph.BuildGraph Application.ActiveWindow.Selection(1)
    
    'Очищаем расчеты узлов графа
    controller.SetGraph(graph).ClearGraph.ShapesRefresh
    controller.SetF 0.1       'Указываем площадь человека в одежде "летняя;весенне-осенняя;зимняя"/"0.1;0.113;0.125"
    controller.SetF InputBox("Площадь проекции?", "Укажите площадь проекции человека", 0.1)
    controller.ResolveGraph_PeopleFlow
    controller.calculate
    controller.ResolveGraph_TimesFlow
    controller.ShapesRefresh
    
    Debug.Print "Общее время эвакуации как сумма времен всех узлов: " & controller.TotalTime
    Debug.Print "Время эвакуации по последнему узлу: " & controller.graph.exitNodes(1).t_flow
'    MsgBox "Общее время эвакуации как сумма времен всех узлов: " & controller.TotalTime & vbNewLine & _
'            "Время эвакуации по последнему узлу: " & controller.graph.exitNodes(1).t_flow
    MsgBox "Время эвакуации: " & Round(controller.graph.exitNodes(1).t_flow, 1) & " мин."
    
    Set graph = Nothing
End Sub


Public Sub SelectNodes()
'Выбираем фигуры узлов графа
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If IsGFSShapeWithIP(shp, indexPers.ipEvacNode) Then
            Application.ActiveWindow.Select shp, visSelect
        End If
    Next shp
End Sub


Public Sub SeekPlace(ShpObj As Visio.Shape)
'Процедура получения параметров фигуры места и присвоение его фигуре узла графа
Dim OtherShape As Visio.Shape
Dim x, y As Double
Dim col As Collection

    On Error GoTo ex
'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'Перебираем все фигуры на странице (Ищем двери)
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
'Перебираем все фигуры на странице (Ищем места)
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
'Находим расстояние до ближайшей стены
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
'Возвращаем ширину помещения равную удвоенному расстоянию до ближайших стен
    GetWidthByWall = GetNearWallDist(shp) * 2
End Function
