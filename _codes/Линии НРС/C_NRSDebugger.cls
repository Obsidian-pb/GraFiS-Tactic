VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_NRSDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------Демон-дебаггер для модели НРС------------------------
Private NRSModel_ As c_NRSModel
Private NRSObservers As Collection


Private Sub Class_Initialize()
    Set NRSObservers = New Collection
End Sub
Private Sub Class_Terminate()
    Set NRSObservers = Nothing
End Sub

Public Function Activate(ByRef nrs_model As c_NRSModel) As C_NRSDebugger
    
    Set NRSModel_ = nrs_model
    
Set Activate = Me
End Function


Public Function PrintModelState() As C_NRSDebugger
Dim node As c_NRSNode
    
    For Each node In NRSModel_.StarterNodes
        PrintNodeState node.ID, 0
    Next node
    
End Function

Public Function PrintModelState2() As C_NRSDebugger
Dim node As c_NRSNode
    
    frmNRSDescription.tbNRSDescription.Text = ""
    
    For Each node In NRSModel_.StarterNodes
        PrintNodeState2 node.ID, 0
    Next node
    
    frmNRSDescription.AddNewString "Общий расход из приборов: " & GetQSum & " л/с"
    frmNRSDescription.AddNewString "------------------------"
    frmNRSDescription.AddNewString "Q - текущий расход, л/с"
    frmNRSDescription.AddNewString "QPrev - расход на предыдущей итерации, л/с"
    frmNRSDescription.AddNewString "hIn - напор на входе в узел, м"
    frmNRSDescription.AddNewString "H - напор генерируемый узлом (например насосом МСП), м"
    frmNRSDescription.AddNewString "hLost - потери напора в узле, м"
    frmNRSDescription.AddNewString "hOut - напор на выходе из узла, м"
    
'    frmNRSDescription.Show
    
End Function

Public Function PrintQSum() As C_NRSDebugger
'Dim node As c_NRSNode
'Dim sum As Single
'
'    For Each node In NRSModel_.EnderNodes
'        sum = sum + node.Q
'    Next node
'Debug.Print "sum(Q)=" & sum

Debug.Print "sum(Q)=" & GetQSum
Set PrintQSum = Me
End Function

Public Function GetQSum() As Single
Dim node As c_NRSNode
Dim sum As Single
    
    For Each node In NRSModel_.EnderNodes
        sum = sum + node.Q
    Next node
    
GetQSum = sum
End Function

Public Sub PrintNodeState(ByRef nodeID As Long, ByVal level As Byte) 'As C_NRSDebugger
Dim node As c_NRSNode
Dim otst As String
Dim i As indexPers
Dim nextNode As c_NRSNode

    Set node = NRSModel_.GetNRSNodeByKey(nodeID)
    
    For i = 0 To level
        otst = otst & vbTab
    Next i
    
    Debug.Print otst & node.ID & "(" & CellVal(node.NodeShape, "User.IndexPers.Prompt", visUnitsString) & _
                "): Q=" & node.Q & vbTab & " QPrev=" & node.PrevQ & vbTab & " hIn=" & node.hIn & vbTab & " H=" & node.h & vbTab & " hLost = " & node.hLost & vbTab & " hOut = " & node.hOut
    
    For Each nextNode In node.OutNodes
        PrintNodeState nextNode.ID, level + 1
    Next nextNode

'Set PrintNodeState = Me
End Sub

Public Sub PrintNodeState2(ByRef nodeID As Long, ByVal level As Byte)
Dim node As c_NRSNode
Dim otst As String
Dim i As indexPers
Dim nextNode As c_NRSNode

    Set node = NRSModel_.GetNRSNodeByKey(nodeID)
    
    For i = 0 To level
        otst = otst & vbTab
    Next i
    
'    frmNRSDescription.tbNRSDescription.Text = frmNRSDescription.tbNRSDescription.Text & _
'        otst & node.ID & "(" & CellVal(node.NodeShape, "User.IndexPers.Prompt", visUnitsString) & _
'        "): Q=" & node.Q & vbTab & "QPrev=" & node.PrevQ & vbTab & "hIn=" & node.hIn & vbTab & "H=" & node.H & vbTab & "hLost = " & node.hLost & vbTab & "hOut = " & node.hOut & vbNewLine
    
    frmNRSDescription.AddNewString otst & node.ID & "(" & CellVal(node.NodeShape, "User.IndexPers.Prompt", visUnitsString) & _
        "): Q=" & node.Q & vbTab & " QPrev=" & node.PrevQ & vbTab & " hIn=" & node.hIn & vbTab & " H=" & node.h & vbTab & " hLost = " & node.hLost & vbTab & " hOut = " & node.hOut
    
    For Each nextNode In node.OutNodes
        PrintNodeState2 nextNode.ID, level + 1
    Next nextNode
    
End Sub


Public Function AddObserver(ByVal nodeID As Long, props As Variant) As C_NRSDebugger
Dim obs As c_NRSNodeObserver
Dim nd As c_NRSNode
Dim i As Integer
    
    Set nd = NRSModel_.GetNRSNodeByKey(nodeID)
    If nd Is Nothing Then
        Set AddObserver = Me
        Exit Function
    End If
    
    Set obs = New c_NRSNodeObserver
    obs.Activate nd
    
    For i = 0 To UBound(props)
        obs.AddProp props(i)
    Next i
    
    AddUniqueCollectionItem NRSObservers, obs
'    NRSObservers.Add obs, CStr(nodeID)
    
Set AddObserver = Me
End Function

Public Function getObserverHistory(ByVal nodeID As Long, prop As NRSProp) As String
Dim obs As c_NRSNodeObserver
    
    On Error GoTo ex
    
    Set obs = GetFromCollection(NRSObservers, CStr(nodeID))

getObserverHistory = obs.GetHistory(prop)
Exit Function
ex:
    getObserverHistory = "err"
End Function
