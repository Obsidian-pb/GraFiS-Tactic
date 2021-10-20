Attribute VB_Name = "agentMove"
Public Enum Directions
    U = 0
    R = 1
    D = 2
    L = 3
End Enum


Public Sub RunTest()
Dim agent As Visio.Shape
Dim step As Integer
'Dim runLen As Integer
Dim dir As Directions
    
Const runLen = 200
    
    
    dir = U
    Set agent = Application.ActivePage.Shapes.ItemFromID(11210)
    
    For step = 0 To runLen
        Select Case dir
            Case Is = Directions.U
                If testWall(agent) Then
                    dir = R
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) - agent.Cells("Height").Result(visNumber)
                Else
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) + agent.Cells("Height").Result(visNumber)
                End If
            Case Is = Directions.R
                If testWall(agent) Then
                    dir = D
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) - agent.Cells("Width").Result(visNumber)
                Else
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) + agent.Cells("Width").Result(visNumber)
                End If
            Case Is = Directions.D
                If testWall(agent) Then
                    dir = L
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) + agent.Cells("Height").Result(visNumber)
                Else
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) - agent.Cells("Height").Result(visNumber)
                End If
            Case Is = Directions.L
                If testWall(agent) Then
                    dir = U
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) + agent.Cells("Width").Result(visNumber)
                Else
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) - agent.Cells("Width").Result(visNumber)
                End If
        End Select
        
        DoEvents
    Next step
    
    
End Sub




Private Function testWall(ByRef agent As Visio.Shape) As Boolean
Dim shpN As Visio.Shape
Dim sel As Visio.Selection

    Set sel = agent.SpatialNeighbors(VisSpatialRelationCodes.visSpatialOverlap, 100, visSpatialFrontToBack)
    
    For Each shpN In sel
'        Debug.Print shpN.Name
        If isWall(shpN) Then
            testWall = True
            Exit Function
        End If
    Next shpN
testWall = False
End Function




Public Sub RunTest2()
Dim agent As Visio.Shape
Dim step As Integer
'Dim runLen As Integer
Dim dir As Directions
    
Const runLen = 200
    
    
    dir = U
    Set agent = Application.ActivePage.Shapes.ItemFromID(11210)
    
    For step = 0 To runLen
        Select Case dir
            Case Is = Directions.U
                If testWall2(agent.Cells("PinX").Result(visNumber), _
                            agent.Cells("PinY").Result(visNumber) + agent.Cells("Height").Result(visNumber), _
                            agent.Cells("Height").Result(visNumber)) Then
                    dir = R
'                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) - agent.Cells("Height").Result(visNumber)
                Else
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) + agent.Cells("Height").Result(visNumber)
                End If
            Case Is = Directions.R
                If testWall2(agent.Cells("PinX").Result(visNumber) + agent.Cells("Width").Result(visNumber), _
                            agent.Cells("PinY").Result(visNumber), _
                            agent.Cells("Height").Result(visNumber)) Then
                    dir = D
'                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) - agent.Cells("Width").Result(visNumber)
                Else
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) + agent.Cells("Width").Result(visNumber)
                End If
            Case Is = Directions.D
                If testWall2(agent.Cells("PinX").Result(visNumber), _
                            agent.Cells("PinY").Result(visNumber) - agent.Cells("Height").Result(visNumber), _
                            agent.Cells("Height").Result(visNumber)) Then
                    dir = L
'                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) + agent.Cells("Height").Result(visNumber)
                Else
                    agent.Cells("PinY").Formula = agent.Cells("PinY").Result(visNumber) - agent.Cells("Height").Result(visNumber)
                End If
            Case Is = Directions.L
                If testWall2(agent.Cells("PinX").Result(visNumber) - agent.Cells("Width").Result(visNumber), _
                            agent.Cells("PinY").Result(visNumber), _
                            agent.Cells("Height").Result(visNumber)) Then
                    dir = U
'                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) + agent.Cells("Width").Result(visNumber)
                Else
                    agent.Cells("PinX").Formula = agent.Cells("PinX").Result(visNumber) - agent.Cells("Width").Result(visNumber)
                End If
        End Select
        
        DoEvents
    Next step
    
    
End Sub

Private Function testWall2(ByVal x As Double, ByVal y As Double, ByVal tolerance As Double) As Boolean
Dim shpN As Visio.Shape
Dim sel As Visio.Selection

    ' Разобраться - скорее всего проблема в единицах измерения x и y
    Set sel = Application.ActivePage.SpatialSearch(x, y, VisSpatialRelationCodes.visSpatialOverlap, 100, visSpatialFrontToBack)
    
    For Each shpN In sel
        Debug.Print shpN.Name
        If isWall(shpN) Then
            testWall2 = True
            Exit Function
        End If
    Next shpN
testWall2 = False
End Function

Public Function isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
Dim shapeType As Integer
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    shapeType = aO_Shape.Cells("User.ShapeType").Result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (shapeType = 44 Or shapeType = 6) Then
        isWall = True
        Exit Function
    End If
isWall = False
End Function
