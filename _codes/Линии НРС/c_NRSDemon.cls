VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------�����-���������� �������-�������� �������-----------------------
Private NRSModel_ As c_NRSModel


Public roundAccuracy As Byte                 '�������� ����������
Public checkAccuracy As Byte                 '�������� ���������� ��� �������� �������������� ������� (� �����)
Public outAccuracy As Byte                   '�������� ���������� ��� ������ ������ � ������
Public maxIterations As Integer              '������������ ���������� �������� �������
Public approvedHout As Single                '���������� ����� �� ������ �� �����

Private calculatedNodes As Collection        '��������� ������������ �����

Private Sub Class_Initialize()
    roundAccuracy = 3
    checkAccuracy = 2
    outAccuracy = 2
    maxIterations = 20
    approvedHout = 10
'        NRSDemon.roundAccuracy = 4
'    NRSDemon.checkAccuracy = 2
'    NRSDemon.outAccuracy = 2
'    NRSDemon.maxIterations = 100
'    NRSDemon.approvedHout = 3
End Sub
Private Sub Class_Terminate()
    Set NRSModel_ = Nothing
End Sub
'Public Function Activate(ByRef shp As Visio.Shape) As c_NRSDemon
'
'Set Activate = Me
'End Function








Public Function ClearNRSModel() As c_NRSDemon
    NRSModel_.Clear
Set ClearNRSModel = Me
End Function
Public Function BuildNRSModel(ByRef shp As Visio.Shape) As c_NRSDemon
'������ ������ �������-�������� ������� �� ���������� ������
Dim modelBuilder As c_NRSModelBuilder

'    On Error GoTo EX
'    ClearNRSModel
    
    Set modelBuilder = New c_NRSModelBuilder
        Set NRSModel_ = modelBuilder.Build(shp)
    Set modelBuilder = Nothing
    
    
    
    
    
    

Set BuildNRSModel = Me
Exit Function
ex:
    Set BuildNRSModel = Me
End Function


Public Function CalculateNRSModel() As c_NRSDemon
'������������ ������ �������-�������� �������: �������� ���������
Dim i As Integer
Dim node As c_NRSNode
'Dim props() As Integer
    
'    On Error GoTo ex
    
    Dim debugger As C_NRSDebugger
    Set debugger = New C_NRSDebugger
    debugger.Activate NRSModel_
    
'    For Each node In NRSModel_.EnderNodes
'        debugger.AddObserver node.ID, Array(NRSProp.nrsPropQ, NRSProp.nrsProphIn)
'    Next
    debugger.AddObserver NRSModel_.StarterNodes(1).ID, Array(NRSProp.nrsPropQ, NRSProp.nrsProphIn)
    

    '���� ��������:
        For i = 1 To maxIterations
        '�������� ��������� ������������ ����� ���
            Set calculatedNodes = New Collection
        '�������� ������� ������ � �����
            ClearhIn
'            debugger.PrintModelState
        '������ �� ��������� � ������������
            CalculateNRSModelForward
'            debugger.PrintModelState

'!����������� ���������� ���������!!!
'��� �������� ������� ������� � ��������� ������� - ��� ������ ���������� ����� �� ��� �� ��������� ������.
'��� �������� ��� ��������� ������, �� �� ��������� �����. ������� ��������� ������� ����� ������! ������������ �� ����� ��� ������ ���������� �� ������ �����!!! 0.000001 >> 0.01

        '��������� ��������� ������
            If Not NRSModel_.correct Then
                Realize.CheckNodes
                Debug.Print "�������� " & i
'                debugger.PrintQSum
'                debugger.PrintModelState
                debugger.PrintModelState2
                MsgBox "������� �������-�������� ������� ���������������� ��� ��������� ���������� ���������!", vbInformation
                Set CalculateNRSModel = Me
                Exit Function
            End If
'            If NRSModel_.TotalQ = 0 And i > 1 Then
'                Realize.CheckNodes
'                Debug.Print "�������� " & i
'                debugger.PrintQSum
'                debugger.PrintModelState
'                MsgBox "������� �������-�������� ������� ���������������� ��� ��������� ���������� ���������!", vbInformation
'                Set CalculateNRSModel = Me
'                Exit Function
'            End If
            
        '�������� ��������� ������������ ����� ���
            Set calculatedNodes = New Collection
        '�������� ������� ������� � �����
            ClearQ
'            debugger.PrintModelState
        '������ �� ������������ � ���������
            CalculateNRSModelBackward
'            debugger.PrintModelState

            debugger.ObsHistoryExpand
            
'            '��������� ������ ������ ��� �����!!!
'            Realize.CheckNodes
            
            If CheckStacionar Then Exit For
        Next i
        
        
        '��������� ���������� ������� � �������
        Realize.CheckNodes
    Debug.Print i
'    debugger.PrintModelState
    debugger.PrintModelState2
    debugger.PrintQSum
'    Stop
'    For Each node In NRSModel_.EnderNodes
'        Debug.Print debugger.getObserverHistory(node.ID, NRSProp.nrsPropQ)
'        Debug.Print debugger.getObserverHistory(node.ID, NRSProp.nrsProphIn)
'    Next
    Debug.Print debugger.getObserverHistory(NRSModel_.StarterNodes(1).ID, NRSProp.nrsPropQ)
    
    Set debugger = Nothing
    
Set CalculateNRSModel = Me
Exit Function
ex:
    Set CalculateNRSModel = Me
End Function


Public Function CalculateNRSModelForward() As c_NRSDemon
'���� ���� ������� ������� ������ �������-�������� �������
Dim node As c_NRSNode

    On Error GoTo ex
    
    For Each node In NRSModel_.StarterNodes
        CalculateNRSNodeForward node
    Next node
    
    
    
Set CalculateNRSModelForward = Me
Exit Function
ex:
    Set CalculateNRSModelForward = Me
End Function

Public Function CalculateNRSNodeForward(ByRef node As c_NRSNode) As c_NRSDemon
'������ ������ ���� �������-�������� �������
Dim nextNode As c_NRSNode
Dim tmpVal As Single
    
'    Debug.Print node.ID
    
    On Error GoTo ex
    
    For Each nextNode In node.OutNodes
'        Debug.Assert nextNode.ID <> 19
'        Debug.Assert nextNode.ID <> 67
'        Debug.Assert node.ID <> 19
'        Debug.Assert node.ID <> 324
        tmpVal = GetHOut(node)
'        Debug.Assert node.ID <> 85
        If tmpVal > nextNode.hIn Then
            Stop    '����� ����� �������� ���� ������������� ����� ���� �������������� � �.�. (nextNode.hIn = ��������*tmpVal) ��� ���� ����� �� ������ �������� �������� ������� ������ ��������������!
            nextNode.hIn = tmpVal     '���������, ��� ����� �� ����� ����� ������������� ������ �� ��������
'            CalculateNRSNodeForward nextNode
            
            If Not IsInCollection(calculatedNodes, nextNode) Then
                AddUniqueCollectionItem calculatedNodes, nextNode
                CalculateNRSNodeForward nextNode
            End If
            
        End If
'        nextNode.hIn = GetHOut(node)    'node.hOut
'        Debug.Assert node.ID <> 324
'        CalculateNRSNodeForward nextNode
    Next nextNode
    
    
    
Set CalculateNRSNodeForward = Me
Exit Function
ex:
    Set CalculateNRSNodeForward = Me
End Function


Public Function ClearQ() As c_NRSDemon
'������� ������� �� ���� ����� ����� Enders
Dim node As c_NRSNode

    On Error GoTo ex

    For Each node In NRSModel_.NRSNodes
'        If Not node.isEnder Then node.Q_Clear
        node.Q_Clear
    Next node
    
Set ClearQ = Me
Exit Function
ex:
    Set ClearQ = Me
End Function

Public Function ClearhIn() As c_NRSDemon
'������� �������� ������ �� ���� ����� ����� starters (���������)
Dim node As c_NRSNode

    On Error GoTo ex

    For Each node In NRSModel_.NRSNodes
'        If Not node.isStarter Then node.hIn_Clear
        node.hIn_Clear
    Next node
    
Set ClearhIn = Me
Exit Function
ex:
    Set ClearhIn = Me
End Function



Public Function CalculateNRSModelBackward() As c_NRSDemon
'���� ���� ��������� ������� ������ �������-�������� �������
Dim node As c_NRSNode

    On Error GoTo ex
       
    For Each node In NRSModel_.EnderNodes
        CalculateNRSNodeBackward node, GetQStv(node) 'node.QStv
    Next node
    
Set CalculateNRSModelBackward = Me
Exit Function
ex:
    Set CalculateNRSModelBackward = Me
End Function

Public Function CalculateNRSNodeBackward(ByRef node As c_NRSNode, ByVal a_Q As Single) As c_NRSDemon
'�������� ������ ���� �������-�������� �������
Dim prevNode As c_NRSNode
Dim newQ As Single
    
    On Error GoTo ex
    
    For Each prevNode In node.InNodes
'        If node.isEnder Then
'            prevNode.Q_Raise node.QStv
'            Debug.Print node.ID & " Q=" & node.QStv & ">>" & prevNode.ID & ": Q=" & prevNode.Q & "+" & node.QStv & "=" & prevNode.Q
'        Else
'            prevNode.Q_Raise node.Q
'            Debug.Print node.ID & " Q=" & node.Q & ">>" & prevNode.ID & ": Q=" & prevNode.Q & "+" & node.Q & "=" & prevNode.Q
'        End If
'        Debug.Assert node.InNodes.Count < 2
'        newQ = RInt(a_Q / node.InNodes.Count, roundAccuracy)
        newQ = Round(a_Q / node.InNodes.Count, roundAccuracy)
        Stop    '����� ����� �������� ���� ������������� ����� ���� �������������� � �.�. (prevNode.Q_Raise =newQ/(1-��������) �� ������ �������� ������� ����������� ������� ���������������
        prevNode.Q_Raise newQ
'        CalculateNRSNodeBackward prevNode, newQ
        
'        Debug.Assert node.ID <> 7
        If Not IsInCollection(calculatedNodes, prevNode) Then
            AddUniqueCollectionItem calculatedNodes, prevNode
            CalculateNRSNodeBackward prevNode, newQ
        End If
    Next prevNode
    
    
    
Set CalculateNRSNodeBackward = Me
Exit Function
ex:
    Set CalculateNRSNodeBackward = Me
End Function


Public Function Realize() As c_NRSDemon
'��������� ���������� ������� � �������
Dim node As c_NRSNode
       
    On Error GoTo ex
    
    For Each node In NRSModel_.NRSNodes
        RealizeNode node
    Next node
    
Set Realize = Me
Exit Function
ex:
    Set Realize = Me
End Function

Private Function CheckStacionar() As Boolean
Dim diff As Double

    diff = Abs(NRSModel_.TotalQ - NRSModel_.TotalQPrev)
    CheckStacionar = diff <= 0.1 ^ checkAccuracy

End Function



'-------------------------���� ������ � ������� � ������------------------------------------------
'----------------------������� ������ � ������� �����---------------------
Private Function GetHOut(ByRef node As c_NRSNode) As Single
Dim val As Single
'    Debug.Assert node.ID <> 85
'    GetHLost node
    val = node.hIn + node.H - GetHLost(node)
'    Debug.Assert node.ID <> 85
    If val < 0 Then
        val = 0
        NRSModel_.correct = False
'        SwithNodeCor node, False
    End If
    If val < approvedHout Then
        NRSModel_.correct = False
'        SwithNodeCor node, False
    End If
    node.hOut = val
    GetHOut = node.hOut
End Function

Private Function GetHLost(ByRef node As c_NRSNode) As Single
'    node.hLost = RInt(node.S * node.Q ^ 2 - node.Z, roundAccuracy)
    node.hLost = Round(node.S * node.Q ^ 2 - node.Z, roundAccuracy)
    GetHLost = node.hLost
End Function

Private Function GetQStv(ByRef node As c_NRSNode) As Single
'������������������ ������� (� �������� ������� ������������)
'    node.Q = RInt(node.P * Sqr(node.hIn - GetHLost(node)), roundAccuracy)
    node.Q = Round(node.P * Sqr(node.hIn - GetHLost(node)), roundAccuracy)
    GetQStv = node.Q
End Function



'Public Function CalcH() As c_NRSDemon
'    curNRSNode_.H = Round(curNRSNode_.S * curNRSNode_.Q ^ 2, roundAccuracy_)
'Set CalcH = Me
'End Function
'
'Public Function CalcQ() As c_NRSModelControllerGidro
'    curNRSNode_.Q = curNRSNode_.Q / curNRSNode_.InNodes.Count
'Set CalcQ = Me
'End Function



'Public Function hLost() As Double
'    hLost = curNRSNode_.S * curNRSNode_.Q ^ 2
'End Function
'
'Public Function hOut() As c_NRSModelControllerGidro
''    hOut = hOut_
''    Debug.Assert nodeShape_.ID <> 324
'    curNRSNode_.hOut = curNRSNode_.H + curNRSNode_.hIn - hLost - Z_
'Set c_NRSModelControllerGidro = Me
'End Function
'
'Public Function Q_Clear() As c_NRSModelControllerGidro
'    curNRSNode_.Q = 0
'Set Q_Clear = Me
'End Function
'Public Function Q_Raise(ByVal a_Q As Single) As c_NRSModelControllerGidro
'    curNRSNode_.Q = curNRSNode_.Q + a_Q
'Set Q_Raise = Me
'End Function
'
'Public Function CalchOut() As c_NRSModelControllerGidro
''    hOut = hOut_
''    ����������� �����!
'    hOut = curNRSNode_.H + curNRSNode_.hIn - curNRSNode_.hLost - curNRSNode_.Z
'Set CalchOut = Me
'End Function
'
'
'
'Private Function CheckStarterEnder() As c_NRSModelControllerGidro
'    curNRSNode_.isStarter = IsGFSShapeWithIP(curNRSNode_.NodeShape, startersArray_)
'    curNRSNode_.isEnder = IsGFSShapeWithIP(curNRSNode_.NodeShape, enderArray_)
'
'Set CheckStarterEnder = Me
'End Function





'---------------------------������� ���������������� ���������---------------------------
Public Function RealizeNode(ByRef node As c_NRSNode) As c_NRSDemon
'����� ���������� �������� ������� �������� ��� ������, �� ������ �������
    On Error GoTo ex
'    Application.EventsEnabled = False
    
    '��� ������ ����� ������
    Select Case CellVal(node.NodeShape, "User.IndexPers")
        Case Is = indexPers.ipAC
            SetCellVal node.NodeShape, "Prop.PodOut", Round(node.Q, outAccuracy)
'            SetCellVal NRSNode.NodeShape, "Prop.PumpPressureCur", Q_
        Case Is = indexPers.ipRukavLineNapor
            SetCellVal node.NodeShape, "Prop.Flow", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "Prop.HeadInHose", Round(node.hOut, outAccuracy)
            SetCellVal node.NodeShape, "Prop.HeadLost", Round(node.hLost, outAccuracy)
        Case Is = indexPers.ipRazvetvlenie
            SetCellVal node.NodeShape, "Prop.StreamThrough", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "Prop.HeadInRazv", Round(node.hOut, outAccuracy)
        Case Is = indexPers.ipStvolRuch
'            SetCellVal node.NodeShape, "User.PodOut", Round(GetQStv(node), outAccuracy)
            SetCellVal node.NodeShape, "User.PodOut", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "Prop.PodOutS", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "User.Head", Round(node.hIn, outAccuracy)
            SetCellVal node.NodeShape, "Prop.HeadS", Round(node.hIn, outAccuracy)
            '!!!�� ������ ��������� ������ ������: =GUARD(IF(Scratch.A1-Prop.StvolHeight<0,0,Scratch.A1-Prop.StvolHeight))
        Case Is = indexPers.ipStvolLafVoda
            SetCellVal node.NodeShape, "User.PodOut", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "Prop.PodOutS", Round(node.Q, outAccuracy)
            SetCellVal node.NodeShape, "User.Head", Round(node.hIn, outAccuracy)
            SetCellVal node.NodeShape, "Prop.HeadS", Round(node.hIn, outAccuracy)
            
    End Select
    
    
'Application.EventsEnabled = True
Set RealizeNode = Me
Exit Function
ex:
'    Application.EventsEnabled = True
    Set RealizeNode = Me
End Function

Public Function CheckNodes() As c_NRSDemon
'����� ���������� �������� ����� �� ������������ ������� � ���� ���� �����������, �� ������ ���� ������������ �������
Dim node As c_NRSNode

    On Error GoTo ex
'    Application.EventsEnabled = False
    
    For Each node In NRSModel_.NRSNodes
        Select Case CellVal(node.NodeShape, "User.IndexPers")
            Case Is = indexPers.ipRukavLineNapor
                If node.hOut = 0 Or node.hOut < approvedHout Then
                    SwithNodeCor node, False
                Else
                    SwithNodeCor node, True
                End If
        End Select
    Next node
    
    
'Application.EventsEnabled = True
Set CheckNodes = Me
Exit Function
ex:
'    Application.EventsEnabled = True
    Set CheckNodes = Me
End Function
Public Function SwithNodeCor(ByRef node As c_NRSNode, ByVal correct As Boolean) As c_NRSDemon
'����� ���������� ������������ ��������� ���� - ���������-���������� (������������ ������������ �������)

    On Error GoTo ex
'    Application.EventsEnabled = False
    
        Select Case CellVal(node.NodeShape, "User.IndexPers")
            Case Is = indexPers.ipRukavLineNapor
                If correct Then
'                    SetCellVal node.NodeShape, "LineColor", "Styles!���_��������!LineColor"
                    SetCellFrml node.NodeShape, "LineColor", "IF(Prop.Flow>Prop.FlowS,2,Styles!�_���!LineColor)"
                Else
                    SetCellFrml node.NodeShape, "LineColor", 2
                End If
        End Select
    
    
'Application.EventsEnabled = True
Set SwithNodeCor = Me
Exit Function
ex:
'    Application.EventsEnabled = True
    Set SwithNodeCor = Me
End Function




