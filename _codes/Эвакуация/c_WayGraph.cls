VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_WayGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------Класс графа путей---------------------
Public nodes As Collection
Public startNodes As Collection
Public exitNodes As Collection









'Private Sub Class_Initialize()
'    Set node = New Collection
'End Sub






Public Function BuildGraph(ByRef shp As Visio.Shape) As c_WayGraph
    Set nodes = New Collection
    Set startNodes = New Collection
    Set exitNodes = New Collection
    
    AddNode NewNode(shp)
    

    
    
    
    
Set BuildGraph = Me
End Function


Public Sub AddNode(ByRef node As c_Node)
Dim link As Visio.Connect
Dim nextShp As Visio.Shape
Dim nextShp2 As Visio.Shape
Dim shp As Visio.Shape
Dim nextNode As c_Node
    
    If IsKeyInCollection(nodes, str(node.number)) Then Exit Sub
    
    Set shp = node.shp
    
    AddUniqueCollectionItem nodes, node, str(node.number)
    If node.nodeType = first Then
        AddUniqueCollectionItem startNodes, node, str(node.number)
    ElseIf node.nodeType = ext Then
        AddUniqueCollectionItem exitNodes, node, str(node.number)
    End If
    
    'Перебираем все исходящие связи
    For Each link In shp.FromConnects
        Set nextShp = link.FromSheet
        
        If nextShp.Connects(1).ToSheet.ID = shp.ID Then
        ' Исходящее соединение
            Set nextShp2 = nextShp.Connects(2).ToSheet
'            Debug.Print "out: " & nextShp2.Name
            Set nextNode = NewNode(nextShp2)
            node.AddNextNode nextNode
            nextNode.AddPrevNode node
        Else
        ' Входящее соединения
            Set nextShp2 = nextShp.Connects(1).ToSheet
'            Debug.Print "in: " & nextShp.Connects(1).ToSheet.Name
            Set nextNode = NewNode(nextShp2)
            node.AddPrevNode nextNode
            nextNode.AddNextNode node
        End If
        
        AddNode nextNode
        
    Next link
    
End Sub














'-------tools---------
Private Function NewNode(ByRef shp As Visio.Shape) As c_Node
'Создаем новый узел

    Set NewNode = New c_Node
    Set NewNode.shp = shp
    NewNode.number = cellVal(shp, "Prop.NodeNumber")
    NewNode.width = cellVal(shp, "Prop.WayWidth")
    NewNode.lenight = cellVal(shp, "Prop.WayLen")
    NewNode.peoplesHere = cellVal(shp, "Prop.PeopleHere")
'    NewNode.q_max = cellval(shp, "Prop.q_max")
    
    'Проверяем к какому классу относится узел (Горизонтальный;Дверной проем;Лестница вниз;Лестница вверх)
    Select Case cellVal(shp, "Prop.WayClass", visUnitsString)
        Case Is = "Горизонтальный"
            NewNode.nodeClass = hor
            NewNode.q_max = 16.5
        Case Is = "Дверной проем"
            NewNode.nodeClass = door
            NewNode.q_max = 19.6
        Case Is = "Лестница вниз"
            NewNode.nodeClass = laddo
            NewNode.q_max = 16
        Case Is = "Лестница вверх"
            NewNode.nodeClass = ladup
            NewNode.q_max = 11
    End Select
    
    'Проверяем тип узла (Начальный;Промежуточный;Выход)
    Select Case cellVal(shp, "Prop.WayType", visUnitsString)
        Case Is = "Начальный"
            NewNode.nodeType = first
        Case Is = "Промежуточный"
            NewNode.nodeType = nest
        Case Is = "Выход"
            NewNode.nodeType = ext
    End Select

End Function
