VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ControllerGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'----------����� ����������� �����-----------------
Public graph As c_WayGraph
Private f As Single

Private vdq_tool As c_VDQ

Private Sub Class_Initialize()
    Set vdq_tool = New c_VDQ
End Sub







Public Function SetGraph(ByRef gph As c_WayGraph) As c_ControllerGraph
'��������� ������ �� ��� � ������� ����� ��������
    Set graph = gph
Set SetGraph = Me
End Function
Public Function SetF(f_a As Single) As c_ControllerGraph
'������������� ������� �������� ������ ��������
    f = f_a
Set SetF = Me
End Function


'---�������� ������� ����������-----------------
Public Function ResolveGraph_PeopleFlow() As c_ControllerGraph
'������ �������� ������ �� ���� ����� �����
Dim node As c_Node
    
    For Each node In graph.startNodes
        node.RaisePeopleFlow node.peoplesHere
        people_calc node, node.peoplesHere
    Next node

Set ResolveGraph_PeopleFlow = Me
End Function
Private Sub people_calc(node As c_Node, val As Integer)
Dim nextNode As c_Node
       
    For Each nextNode In node.nextNodes
        nextNode.RaisePeopleFlow val
        people_calc nextNode, val
    Next nextNode
End Sub

Public Function ResolveGraph_TimesFlow() As c_ControllerGraph
'������ ������ ��������� ��� ���� ����� �����
Dim node As c_Node

    For Each node In graph.nodes
'        If node.nodeType = first Or node.nodeType = nest Then
            node.RaiseTimeFlow node.t_here
            times_calc node, node.t_here
'        End If

    Next node

Set ResolveGraph_TimesFlow = Me
End Function
Private Sub times_calc(node As c_Node, val As Single)
Dim nextNode As c_Node
       
    For Each nextNode In node.nextNodes
        nextNode.RaiseTimeFlow val
        times_calc nextNode, val
    Next nextNode
End Sub








Public Function calculate() As c_ControllerGraph
'������
Dim node As c_Node
    
    For Each node In graph.startNodes
        CalcNode_Start node
    Next node
    
    
Set calculate = Me
End Function


Private Sub CalcNode_Start(ByRef node As c_Node)
'������ ���������� ��� ���������� ����
Dim nextNode As c_Node
    
    On Error GoTo ex
    
    node.D = (node.peoplesFlow * f) / (node.lenight * node.width)
    node.q = vdq_tool.GetQ(node.D, node.nodeClass)
    node.V = vdq_tool.GetV(node.D, node.nodeClass)
    
    If node.V > 0 Then
        node.t_here = Round(node.lenight / node.V, 2)
    Else
        node.t_here = 0
    End If
    
    '���������� ��� ��������� ����
    For Each nextNode In node.nextNodes
        '� ������, ���� ��������� ���� �� �������� ���������, ����������� ��� ���� ���������� �� �������� ����
        If nextNode.nodeType = nest Then
            CalcNode_Path nextNode
        ElseIf nextNode.nodeType = ext Then
            CalcNode_Exit nextNode
        End If
    Next nextNode
Exit Sub
ex:
    MsgBox "�� ������ �������� ���������� � ���� � ������� " & node.number
End Sub


Private Sub CalcNode_Path(ByRef node As c_Node)
'������ ���������� ��� �������������� ����
Dim nextNode As c_Node
    
    '���������� �������������
    node.q = node.qw_prevs / node.width
    
    If node.q <= node.q_max Then
    '�������� ���
        Select Case node.nodeClass
            Case Is = wayClasses.door '������� �����
                node.t_here = 0
            Case Else
                node.V = vdq_tool.GetVByQ(node.q, node.nodeClass)
                node.t_here = Round(node.lenight / node.V, 2)
        End Select
    Else
    '���� ��������
        node.D = 0.9
        Select Case node.nodeClass
            Case Is = wayClasses.door '������� �����
                '��������� ��� ������� ������� �������� �� ��������, �� �������� ������� ������ �� ������ ��������������� �������
                node.V = vdq_tool.GetVByQ(node.q, wayClasses.hor)
                
'                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.nextNodes(1).width)) - (1 / (node.qw)))
                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw_prevs)))
                node.t_sk = node.peoplesFlow * f / (vdq_tool.GetQ(node.D, node.nodeClass) * node.nextNodes(1).width)
                node.t_here = Round((node.lenight / node.V) + node.t_zad, 2)
            Case Else
                node.V = vdq_tool.GetVByQ(node.q, node.nodeClass)
                
'                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.nextNodes(1).width)) - (1 / (node.qw)))
                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw_prevs)))
                node.t_sk = node.peoplesFlow * f / (vdq_tool.GetQ(node.D, node.nodeClass) * node.nextNodes(1).width)
                node.t_here = Round((node.lenight / node.V) + node.t_zad, 2)
        End Select
    End If
    
    '���������� ��� ��������� ����
    For Each nextNode In node.nextNodes
        '� ������, ���� ��������� ���� �� �������� ���������, ����������� ��� ���� ���������� �� �������� ����
        If nextNode.nodeType = nest Then
            CalcNode_Path nextNode
        ElseIf nextNode.nodeType = ext Then
            CalcNode_Exit nextNode
        End If
    Next nextNode
End Sub

Private Sub CalcNode_Exit(ByRef node As c_Node)
'������ ���������� ��� ���� ������
Dim nextNode As c_Node
    
    '���������� �������������
    node.q = node.qw_prevs / node.width
    
    If node.q <= node.q_max Then
    '�������� ���
        Select Case node.nodeClass
            Case Is = wayClasses.door '������� �����
                node.t_here = 0
            Case Else
                node.V = vdq_tool.GetVByQ(node.q, node.nodeClass)
                node.t_here = Round(node.lenight / node.V, 2)
        End Select
    Else
    '���� ��������
        node.D = 0.9
        Select Case node.nodeClass
            Case Is = wayClasses.door '������� �����
                '��������� ��� ������� ������� �������� �� ��������, �� �������� ������� ������ �� ������ ��������������� �������
                node.V = vdq_tool.GetVByQ(node.q, wayClasses.hor)
                
'                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw)))
                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw_prevs)))
                node.t_sk = node.peoplesFlow * f / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)
                node.t_here = Round((node.lenight / node.V) + node.t_zad, 2)
            Case Else
                node.V = vdq_tool.GetVByQ(node.q, node.nodeClass)
                
'                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw)))
                node.t_zad = node.peoplesFlow * f * ((1 / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)) - (1 / (node.qw_prevs)))
                node.t_sk = node.peoplesFlow * f / (vdq_tool.GetQ(node.D, node.nodeClass) * node.width)
                node.t_here = Round((node.lenight / node.V) + node.t_zad, 2)
        End Select
    End If
    
End Sub



'--------------------��������----------------------
Public Property Get TotalTime() As Variant
Dim node As c_Node
    
    For Each node In graph.nodes
        TotalTime = TotalTime + node.t_here
    Next node
End Property

'--------------------�����������-------------------
Public Function ClearGraph() As c_ControllerGraph
'������� ���������� ������� � ����� �����
Dim node As c_Node
    
    For Each node In graph.nodes
        node.ClearCalculatedData
    Next node
    
Set ClearGraph = Me
End Function

Public Function ShapesRefresh() As c_ControllerGraph
'��������� �������� � ������� �����
Dim node As c_Node
    
    For Each node In graph.nodes
        SetCellVal node.shp, "Prop.tHere", node.t_here
        SetCellVal node.shp, "Prop.t_Flow", node.t_flow
        SetCellVal node.shp, "Prop.PeopleFlow", node.peoplesFlow
        
        SetCellVal node.shp, "Prop.D", node.D
        SetCellVal node.shp, "Prop.V", node.V
        SetCellVal node.shp, "Prop.q", node.q
        SetCellVal node.shp, "Prop.qw", node.qw
        
    Next node

Set ShapesRefresh = Me
End Function

Public Function RenumNodes() As c_ControllerGraph
'���������������� ���� �� �������
Dim node As c_Node
Dim i As Integer
    
    For Each node In graph.nodes
        i = i + 1
        
        node.number = i
        SetCellVal node.shp, "Prop.NodeNumber", i
    Next node
End Function






