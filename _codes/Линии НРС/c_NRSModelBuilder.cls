VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSModelBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------Демон-строитель насосно-рукавной системы-----------------------
Private NRSModel_ As c_NRSModel



Public Function Build(ByRef shp As Visio.Shape) As c_NRSModel
Dim node As c_NRSNode
    
    Set NRSModel_ = New c_NRSModel
    
    Set node = New c_NRSNode
    node.Activate shp
    NRSModel_.AddNRSNode node
    
    GetTechShapeForGESystem node
    
'    FillStartersEnders
    
'    Debug.Print NRSModel_.GetNodesCount
    
Set Build = NRSModel_
End Function



'Private Sub GetTechShapeForGESystem(ByRef shp As Visio.Shape)
Private Sub GetTechShapeForGESystem(ByRef curNRSNode As c_NRSNode)
'Заполняем коллекцию фигур соединенных в НРС
Dim shp As Visio.Shape
Dim con As Connect
'Dim sideShp As Visio.Shape
'Dim curNRSNode As c_NRSNode
Dim secondNRSNode As c_NRSNode
Dim i As Integer
Dim connPointsCount As Integer
Dim cll As Visio.cell

'    For Each con In shp.Connects
'        If Not IsShapeAllreadyChecked(con.ToSheet) Then
'            shapesInNRS.Add con.ToSheet
'            GetTechShapeForGESystem con.ToSheet
'        End If
'    Next con
'    For Each con In shp.FromConnects
'        If Not IsShapeAllreadyChecked(con.FromSheet) Then
'            shapesInNRS.Add con.FromSheet
'            GetTechShapeForGESystem con.FromSheet
'        End If
'    Next con

    Set shp = curNRSNode.NodeShape
'    Set curNRSNode = NewNRSNode(shp)
'    Debug.Assert shp.ID <> 19
    
    If IsGFSShapeWithIP(shp, indexPers.ipRukavLineNapor) Then       'Если текущая фигура рукав, то проверяем подключения к его концам...
        For Each con In shp.Connects
            If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
                Set secondNRSNode = New c_NRSNode
                secondNRSNode.Activate con.ToSheet
'                NRSModel_.AddNRSNode secondNRSNode
'                Debug.Assert secondNRSNode.NodeShape.ID <> 19
                If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
                    curNRSNode.AddOutNode secondNRSNode
                        secondNRSNode.AddInNode curNRSNode
'                        GetTechShapeForGESystem secondNRSNode
                ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
                    curNRSNode.AddInNode secondNRSNode
                        secondNRSNode.AddOutNode curNRSNode
'                        GetTechShapeForGESystem secondNRSNode
                Else
                    'ЗДЕСЬ ДЛЯ ДРУГИХ РУКАВНЫХ ЛИНИЙ
                End If
                
'                secondNRSNode = NewNRSNode(con.ToSheet)
                If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
''                    Debug.Assert secondNRSNode.NodeShape.ID <> 19
'                    If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
'                        curNRSNode.AddOutNode secondNRSNode
''                        GetTechShapeForGESystem secondNRSNode
'                    ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
'                        curNRSNode.AddInNode secondNRSNode
''                        GetTechShapeForGESystem secondNRSNode
'                    Else
'                        'ЗДЕСЬ ДЛЯ ДРУГИХ РУКАВНЫХ ЛИНИЙ
'                    End If
                    NRSModel_.AddNRSNode secondNRSNode
                    GetTechShapeForGESystem secondNRSNode
                End If
'                NRSModel_.AddNRSNode secondNRSNode
'                GetTechShapeForGESystem secondNRSNode
'                Debug.Print con.FromCell.Name
            End If
        Next con
    Else                                                            '...иначе проверяем подключения к точкам соединения
        connPointsCount = shp.RowCount(visSectionConnectionPts)
        If connPointsCount > 0 Then
'            If Left(cll.Name, 18) = "Connections.GFS_Ou" Then
            For i = 0 To connPointsCount - 1
                Set cll = shp.CellsSRC(visSectionConnectionPts, i, 0)
                If Left(cll.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
                            Set secondNRSNode = New c_NRSNode
                            secondNRSNode.Activate con.FromSheet
                            curNRSNode.AddOutNode secondNRSNode
                                secondNRSNode.AddInNode curNRSNode
'                            NRSModel_.AddNRSNode secondNRSNode
                
'                            secondNRSNode = NewNRSNode(con.ToSheet)
                            If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
'                                Debug.Assert secondNRSNode.NodeShape.ID <> 19
'                                curNRSNode.AddOutNode secondNRSNode
                                NRSModel_.AddNRSNode secondNRSNode
                                GetTechShapeForGESystem secondNRSNode
'                                NRSModel_.AddNRSNode secondNRSNode
                            End If
                            
'                            Debug.Print cll.Name
                        End If
                    Next con
                ElseIf Left(cll.Name, Len(ccs_InIdent)) = ccs_InIdent Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
                            Set secondNRSNode = New c_NRSNode
                            secondNRSNode.Activate con.FromSheet
                            curNRSNode.AddInNode secondNRSNode
                                secondNRSNode.AddOutNode curNRSNode
'                            NRSModel_.AddNRSNode secondNRSNode
                            
'                            secondNRSNode = NewNRSNode(con.FromSheet)
                            If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
'                                Debug.Assert secondNRSNode.NodeShape.ID <> 19
'                                curNRSNode.AddInNode secondNRSNode
                                NRSModel_.AddNRSNode secondNRSNode
                                GetTechShapeForGESystem secondNRSNode
'                                NRSModel_.AddNRSNode secondNRSNode
                            End If
                            
'                            Debug.Print cll.Name
                        End If
                    Next con
                End If
            Next i
'            ElseIf Left(cll.Name, 20) = "Connections.GFS_In" Then
'
'            End If
        End If

        
        
    End If
    
    
    
    
End Sub

'Private Sub FillStartersEnders()
'Dim node As c_NRSNode
'
'    For Each node In NRSModel_.NRSNodes
'        If node.isStarter Then NRSModel_.AddStarter node
'    Next node
'End Sub











Private Function NewNRSNode(ByRef shp As Visio.Shape) As c_NRSNode
Dim node As c_NRSNode
    
    Set node = New c_NRSNode
    node.Activate shp
    
    Set NewNRSNode = node       'Потом добавить функцию определения расходов и сопротивлений!
End Function
