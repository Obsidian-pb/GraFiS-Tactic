Attribute VB_Name = "m_tests"
Option Explicit




'Public Sub TTT()
'Dim graph As c_WayGraph
'Dim controller As c_ControllerGraph
'
'    Set graph = New c_WayGraph
'    Set controller = New c_ControllerGraph
'
'    'Строим граф
'    graph.BuildGraph Application.ActiveWindow.Selection(1)
'
'    'Очищаем расчеты узлов графа
'    controller.SetGraph(graph).ClearGraph.ShapesRefresh
'    controller.SetF 0.1       'Указываем площадь человека в одежде "летняя;весенне-осенняя;зимняя"/"0.1;0.113;0.125"
'    controller.ResolveGraph_PeopleFlow.calculate.ResolveGraph_TimesFlow.ShapesRefresh
'
'    Debug.Print "Общее время эвакуации как сумма времен всех узлов: " & controller.TotalTime
'    Debug.Print "Время эвакуации по последнему узлу: " & controller.graph.exitNodes(1).t_flow
'
'    Set graph = Nothing
'End Sub
