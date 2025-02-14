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
'-------------------�����-��������� �������-�������� �������-----------------------
Private NRSModel_ As c_NRSModel

Private startersArray_ As Variant
Private enderArray_ As Variant


Private Sub Class_Initialize()
    startersArray_ = Array(ipAC, ipANR, ipPNS, ipAA, ipAV, ipAKT, ipAP, ipAGVT, ipPoezd, ipMotoPump, ipPricep, ipKorabl, ipKater, ipACL, ipACKP, ipAPP, ipKolonka)
    enderArray_ = Array(ipStvolRuch, ipStvolRuchPena, ipStvolLafVoda, ipStvolLafPena, ipStvolLafPoroshok, ipStvolLafVozimiy, ipPenopodemnik, ipStvolGas, ipStvolPoroshok)
End Sub
Private Sub Class_Terminate()
    Set startersArray_ = Nothing
    Set enderArray_ = Nothing
    Set NRSModel_ = Nothing
End Sub

Public Function Build(ByRef shp As Visio.Shape) As c_NRSModel
Dim node As c_NRSNode
    
    Set NRSModel_ = New c_NRSModel
    
    Set node = New c_NRSNode
'    node.Activate shp
    CheckStarterEnder node.Activate(shp)
    FillProps node
    NRSModel_.AddNRSNode node
    
    GetTechShapeForGESystem node
    
    LinkNRSNodes
    
'    FillStartersEnders
    
'    Debug.Print NRSModel_.GetNodesCount
    
Set Build = NRSModel_
End Function




Private Sub GetTechShapeForGESystem(ByRef curNRSNode As c_NRSNode)
'��������� ��������� ����� ����������� � ���
Dim shp As Visio.Shape
Dim con As Connect
'Dim sideShp As Visio.Shape
'Dim curNRSNode As c_NRSNode
Dim secondNRSNode As c_NRSNode
Dim i As Integer
Dim connPointsCount As Integer
Dim cll As Visio.cell



    Set shp = curNRSNode.NodeShape
'    Debug.Assert shp.ID <> 19
    
    If IsGFSShapeWithIP(shp, indexPers.ipRukavLineNapor) Then       '���� ������� ������ �����, �� ��������� ����������� � ��� ������...
'    Connects:     Sheet.34      Sheet.34      BeginX        Sheet.33      Connections.X1
'    Connects:     Sheet.33      Sheet.33      BeginX        ��            Connections.GFS_OutPatrRight.X
'    FromConnects: Sheet.33      Sheet.34      BeginX        Sheet.33      Connections.X1
'    Connects:     Sheet.34      Sheet.34      EndX          ������� ������ �����        Connections.GFS_In.X
'    ===========33:
'    Connects:     Sheet.33      Sheet.33      BeginX        ��            Connections.GFS_OutPatrRight.X
'    FromConnects: Sheet.33      Sheet.34      BeginX        Sheet.33      Connections.X1
        For Each con In shp.Connects
'            Debug.Print "Connects: ", shp.Name, con.FromSheet.Name, con.FromCell.Name, con.ToSheet.Name, con.ToCell.Name
            If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
'                Set secondNRSNode = New c_NRSNode      '������
'                CheckStarterEnder secondNRSNode.Activate(con.ToSheet)
'                FillProps secondNRSNode
'                If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
''                    Debug.Assert shp.ID <> 221
'                    curNRSNode.AddInNode secondNRSNode
'                        secondNRSNode.AddOutNode curNRSNode
'
'                ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
''                    Debug.Assert shp.ID <> 221
'                    curNRSNode.AddOutNode secondNRSNode
'                        secondNRSNode.AddInNode curNRSNode
'                Else
'                    '����� ��� ������ �������� �����
'                End If

                If Not NRSModel_.InModel(con.ToSheet) Then
                    Set secondNRSNode = New c_NRSNode      '������
                    CheckStarterEnder secondNRSNode.Activate(con.ToSheet)
                    FillProps secondNRSNode
                    NRSModel_.AddNRSNode secondNRSNode
                    GetTechShapeForGESystem secondNRSNode
                End If
            End If
            '��� ������ ������� ��� ���������������� ����������
'            If con.ToCell.Name = "Connections.X1" Then
'                Debug.Print con.ToSheet.Name
'                If Not NRSModel_.InModel(con.ToSheet) Then
'                    Set secondNRSNode = New c_NRSNode      '������
'                    CheckStarterEnder secondNRSNode.Activate(con.ToSheet)
'                    FillProps secondNRSNode
'                    NRSModel_.AddNRSNode secondNRSNode
'                    GetTechShapeForGESystem secondNRSNode
'                End If
'            End If
        Next con
        '��� ������ ������� ��� ���������������� ����������
        For Each con In shp.FromConnects
'            Debug.Print "FromConnects: ", shp.Name, con.FromSheet.Name, con.FromCell.Name, con.ToSheet.Name, con.ToCell.Name
            If Not NRSModel_.InModel(con.FromSheet) Then
                Set secondNRSNode = New c_NRSNode      '������
                CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
                FillProps secondNRSNode
                NRSModel_.AddNRSNode secondNRSNode
                GetTechShapeForGESystem secondNRSNode
            End If
        Next con
    Else                                                            '...����� ��������� ����������� � ������ ����������
        connPointsCount = shp.RowCount(visSectionConnectionPts)
        If connPointsCount > 0 Then
            For i = 0 To connPointsCount - 1
                Set cll = shp.CellsSRC(visSectionConnectionPts, i, 0)
                If Left(cll.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
'                            Debug.Assert shp.ID <> 221
'                            Set secondNRSNode = New c_NRSNode      '������
'                            CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
'                            FillProps secondNRSNode
                
                            If Not NRSModel_.InModel(con.FromSheet) Then
                                Set secondNRSNode = New c_NRSNode      '������
                                CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
                                FillProps secondNRSNode
                                NRSModel_.AddNRSNode secondNRSNode
                                GetTechShapeForGESystem secondNRSNode
                            End If
                        End If
                    Next con
                ElseIf Left(cll.Name, Len(ccs_InIdent)) = ccs_InIdent Then
                    For Each con In shp.FromConnects
                        If con.ToCell.Name = cll.Name Then
'                            Debug.Assert shp.ID <> 1
'                            Set secondNRSNode = New c_NRSNode
'                            CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
'                            FillProps secondNRSNode
                            If Not NRSModel_.InModel(con.FromSheet) Then
                                Set secondNRSNode = New c_NRSNode
                                CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
                                FillProps secondNRSNode
                                NRSModel_.AddNRSNode secondNRSNode
                                GetTechShapeForGESystem secondNRSNode
                            End If
                        End If
                    Next con
                End If
            Next i
        End If

        
        
    End If
    
    
    
    
End Sub

Private Sub LinkNRSNodes()
'������� ����� ��� ���� ����� ������ '  � ��������� �� �� �������/�����
Dim curNRSNode As c_NRSNode
Dim shp As Visio.Shape
Dim con As Connect
Dim secondNRSNode As c_NRSNode
Dim i As Integer
Dim connPointsCount As Integer
Dim cll As Visio.cell
Dim flowDir As Integer
    
    For Each curNRSNode In NRSModel_.NRSNodes
        Set shp = curNRSNode.NodeShape
        
        If IsGFSShapeWithIP(shp, indexPers.ipRukavLineNapor) Then       '���� ������� ������ �����, �� ��������� ����������� � ��� ������...
            For Each con In shp.Connects
                If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
    '                Set secondNRSNode = New c_NRSNode      '������
    '                CheckStarterEnder secondNRSNode.Activate(con.ToSheet)
                    Set secondNRSNode = NRSModel_.GetNRSNodeByShape(con.ToSheet)
'                    CheckStarterEnder secondNRSNode
'                    If secondNRSNode Is Nothing Then
'                        Set secondNRSNode = New c_NRSNode
'                        CheckStarterEnder secondNRSNode.Activate(con.ToSheet)
'                    End If
'                    FillProps secondNRSNode
                    If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
    '                    Debug.Assert shp.ID <> 221
                        curNRSNode.AddInNode secondNRSNode
                            secondNRSNode.AddOutNode curNRSNode
                            
                    ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
    '                    Debug.Assert shp.ID <> 221
                        curNRSNode.AddOutNode secondNRSNode
                            secondNRSNode.AddInNode curNRSNode
                    End If
'                    Stop
                    '���� ��� ������ - �������� �����, ��������� ����������� ������� ����������� ����� � ����� �������� � ����� ����������� �������� ����
                    If IsGFSShapeWithIP(con.ToSheet, indexPers.ipRukavLineNapor) Then
                        flowDir = IsHoseForOut(con.ToSheet, shp.ID)
                        If flowDir = -1 Then
'                            curNRSNode.AddOutNode secondNRSNode
'                                secondNRSNode.AddInNode curNRSNode
                            curNRSNode.AddInNode secondNRSNode
                                secondNRSNode.AddOutNode curNRSNode
                        ElseIf flowDir = 1 Then
'                            curNRSNode.AddInNode secondNRSNode
'                                secondNRSNode.AddOutNode curNRSNode
                            curNRSNode.AddOutNode secondNRSNode
                                secondNRSNode.AddInNode curNRSNode
                        End If
                    End If
    
    
'                    If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
'                        NRSModel_.AddNRSNode secondNRSNode
'                        GetTechShapeForGESystem secondNRSNode
'                    End If
                End If
            Next con
            For Each con In shp.FromConnects
                If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
                    Set secondNRSNode = NRSModel_.GetNRSNodeByShape(con.FromSheet)
                    If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
    '                    Debug.Assert shp.ID <> 221
                        curNRSNode.AddInNode secondNRSNode
                            secondNRSNode.AddOutNode curNRSNode
                    ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
    '                    Debug.Assert shp.ID <> 221
                        curNRSNode.AddOutNode secondNRSNode
                            secondNRSNode.AddInNode curNRSNode
                    End If
'                    Stop    '� ���� ����� ���������� �������� ����������� ��������� ������������ ������ �������� ����� � ����������� �� ����������� �������� ������
                    '���� ��� ������ - �������� �����, ��������� ����������� ������� ����������� ����� � ����� �������� � ����� ����������� �������� ����
                    If IsGFSShapeWithIP(con.FromSheet, indexPers.ipRukavLineNapor) Then
                        flowDir = IsHoseForOut(con.FromSheet, shp.ID)
                        If flowDir = -1 Then
'                            curNRSNode.AddOutNode secondNRSNode
'                                secondNRSNode.AddInNode curNRSNode
                            curNRSNode.AddInNode secondNRSNode
                                secondNRSNode.AddOutNode curNRSNode
                        ElseIf flowDir = 1 Then
'                            curNRSNode.AddInNode secondNRSNode
'                                secondNRSNode.AddOutNode curNRSNode
                            curNRSNode.AddOutNode secondNRSNode
                                secondNRSNode.AddInNode curNRSNode
                        End If
                    End If
                End If
                
            Next con
        Else                                                            '...����� ��������� ����������� � ������ ����������
            connPointsCount = shp.RowCount(visSectionConnectionPts)
            If connPointsCount > 0 Then
                For i = 0 To connPointsCount - 1
                    Set cll = shp.CellsSRC(visSectionConnectionPts, i, 0)
                    If Left(cll.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
                        For Each con In shp.FromConnects
                            If con.ToCell.Name = cll.Name Then
    '                            Debug.Assert shp.ID <> 221
    '                            Set secondNRSNode = New c_NRSNode      '������
    '                            CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
                                Set secondNRSNode = NRSModel_.GetNRSNodeByShape(con.FromSheet)
'                                If secondNRSNode Is Nothing Then
'                                    Set secondNRSNode = New c_NRSNode
'                                    CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
'                                End If
'                                FillProps secondNRSNode
                                curNRSNode.AddOutNode secondNRSNode
                                    secondNRSNode.AddInNode curNRSNode
                    
'                                If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
'                                    NRSModel_.AddNRSNode secondNRSNode
'                                    GetTechShapeForGESystem secondNRSNode
'                                End If
                            End If
                        Next con
                    ElseIf Left(cll.Name, Len(ccs_InIdent)) = ccs_InIdent Then
                        For Each con In shp.FromConnects
                            If con.ToCell.Name = cll.Name Then
    '                            Debug.Assert shp.ID <> 1
                                Set secondNRSNode = NRSModel_.GetNRSNodeByShape(con.FromSheet)
'                                Set secondNRSNode = New c_NRSNode
'                                CheckStarterEnder secondNRSNode.Activate(con.FromSheet)
'                                FillProps secondNRSNode
                                curNRSNode.AddInNode secondNRSNode
                                    secondNRSNode.AddOutNode curNRSNode
'                                If Not NRSModel_.InModel(secondNRSNode.NodeShape) Then
'                                    NRSModel_.AddNRSNode secondNRSNode
'                                    GetTechShapeForGESystem secondNRSNode
'                                End If
                            End If
                        Next con
                    End If
                Next i
            End If
    
            
            
        End If

    Next curNRSNode
'Stop
End Sub


Private Function IsHoseForOut(ByRef shp As Visio.Shape, ByVal PrevShpID As Long) As Integer
'��������� �������� �� ������ ������ ������� ����������� �����
'-1 = ������ �������� �����
' 0 = ������ �� �� �� ������� ������ ������ - ������ �� ����������
' 1 = ������ ��������� �����
Dim con As Connect
Dim nextShp As Visio.Shape
Dim flowDir As Integer
    
    If isAlreadyLinked(shp.ID, PrevShpID) Then
        IsHoseForOut = 0
        Exit Function
    End If
    
    For Each con In shp.Connects
        Set nextShp = con.ToSheet
        If nextShp.ID <> PrevShpID Then
        
            If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
                If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
'                    IsHoseForOut = False
                    IsHoseForOut = -1
                    Exit Function
                ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
'                    IsHoseForOut = True
                    IsHoseForOut = 1
                    Exit Function
                End If
                
                If IsGFSShapeWithIP(nextShp, indexPers.ipRukavLineNapor) Then
                    flowDir = IsHoseForOut(nextShp, shp.ID)
                    IsHoseForOut = flowDir
                    If flowDir <> 0 Then Exit Function
                End If
                
            End If
        Else
            IsHoseForOut = 0
        End If
    Next con
    For Each con In shp.FromConnects
        Set nextShp = con.FromSheet
        If nextShp.ID <> PrevShpID Then
            
            If con.FromCell.Name = "BeginX" Or con.FromCell.Name = "EndX" Then
                If Left(con.ToCell.Name, Len(ccs_OutIdent)) = ccs_OutIdent Then
'                    IsHoseForOut = False
                    IsHoseForOut = -1
                    Exit Function
                        
                ElseIf Left(con.ToCell.Name, Len(ccs_InIdent)) = ccs_InIdent Then
'                    IsHoseForOut = True
                    IsHoseForOut = 1
                    Exit Function
                End If
'                Stop    '� ���� ����� ���������� �������� ����������� ��������� ������������ ������ �������� ����� � ����������� �� ����������� �������� ������
                '���� ��� ������ - �������� �����, ��������� ����������� ������� ����������� ����� � ����� �������� � ����� ����������� �������� ����
                If IsGFSShapeWithIP(nextShp, indexPers.ipRukavLineNapor) Then
'                    If IsHoseForOut(nextShp) Then
'                        curNRSNode.AddInNode secondNRSNode
'                            secondNRSNode.AddOutNode curNRSNode
'                    Else
'                        curNRSNode.AddOutNode secondNRSNode
'                            secondNRSNode.AddInNode curNRSNode
'                    End If
'                    IsHoseForOut = IsHoseForOut(nextShp, shp.ID)
                    flowDir = IsHoseForOut(nextShp, shp.ID)
                    IsHoseForOut = flowDir
                    If flowDir <> 0 Then Exit Function
                End If
            End If
            
        Else
            IsHoseForOut = 0
            
        End If
    Next con
End Function

Private Function isAlreadyLinked(ByVal key1 As String, ByVal key2 As String) As Boolean
'��������� �� ������� �� ��� ���� ��� � ����������� ID ������
Dim node As c_NRSNode
Dim nodetmp As c_NRSNode
    
    Set node = NRSModel_.GetNRSNodeByKey(key1)
    
    For Each nodetmp In node.OutNodes
        If nodetmp.ID = key2 Then
            isAlreadyLinked = True
            Exit Function
        End If
    Next nodetmp
    For Each nodetmp In node.InNodes
        If nodetmp.ID = key2 Then
            isAlreadyLinked = True
            Exit Function
        End If
    Next nodetmp
    
isAlreadyLinked = False
End Function


'Private Sub FillStartersEnders()
'Dim node As c_NRSNode
'
'    For Each node In NRSModel_.NRSNodes
'        CheckStarterEnder node
'        If node.isStarter Then NRSModel_.AddStarter node
'    Next node
'End Sub

Private Function CheckStarterEnder(ByRef node As c_NRSNode) As c_NRSModelBuilder
    node.isStarter = IsGFSShapeWithIP(node.NodeShape, startersArray_)
    node.isEnder = IsGFSShapeWithIP(node.NodeShape, enderArray_)

Set CheckStarterEnder = Me
End Function

Private Function FillProps(ByRef node As c_NRSNode) As c_NRSModelBuilder
'��������� �������� ����� � ����������� � �� IndexPers
Dim val1 As Variant
Dim val2 As Variant

    node.P = CellVal(node.NodeShape, "User.ProvKoeff", , 1)

    If CellVal(node.NodeShape, "Prop.ShowLenightDirect", visUnitsString) = "��" Then
'            (Prop.HoseResistance                            / 20)                                         * Prop.LineLenightS                            '* Prop.Flow^2                            + Prop.Different
        node.S = (CellVal(node.NodeShape, "Prop.HoseResistance") / CellVal(node.NodeShape, "User.HoseLength")) * CellVal(node.NodeShape, "Prop.LineLenightS") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    ElseIf CellVal(node.NodeShape, "Prop.ShowLenightDirect", visUnitsString) = "���" Then
'            Prop.HoseResistance                            * Prop.HoseNeed                            '* Prop.Flow^2                            + Prop.Different
        node.S = CellVal(node.NodeShape, "Prop.HoseResistance") * CellVal(node.NodeShape, "Prop.HoseNeed") '* (CellVal(nodeShape_, "Prop.Flow") ^ 2) + CellVal(nodeShape_, "Prop.Different")
    End If
    
    If node.isEnder Then
        node.hIn = CellVal(node.NodeShape, "Scratch.A1")
    ElseIf node.isStarter Then
        val1 = CellVal(node.NodeShape, "Prop.PodOut")
        val2 = CellVal(node.NodeShape, "Prop.Flow")
        If val1 > val2 Then node.Q = val1 Else node.Q = val2
        
        node.H = CellVal(node.NodeShape, "Prop.PumpPressureCur")
    End If

    
    
    
    node.Z = CellVal(node.NodeShape, "Prop.Different") + CellVal(node.NodeShape, "Prop.StvolHeight")
    
Set FillProps = Me
End Function







Private Function NewNRSNode(ByRef shp As Visio.Shape) As c_NRSNode
Dim node As c_NRSNode
    
    Set node = New c_NRSNode
    node.Activate shp
    
    Set NewNRSNode = node       '����� �������� ������� ����������� �������� � �������������!
End Function
