VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------Класс узла графа путей эвакуации-----------
Public number As Integer
Public width As Single
Public lenight As Single
Public peoplesHere As Integer

Public nodeClass As wayClasses
Public nodeType As wayTypes
Public q_max As Single

Public shp As Visio.Shape
Public nextNodes As Collection
Public prevNodes As Collection

Public peoplesFlow As Integer
Public t_here As Single
Public t_flow As Single
Public t_zad As Single
Public t_sk As Single

Public D As Single
Public V As Single
Public q As Single
'Public qw As Single



Private Sub Class_Initialize()
    Set nextNodes = New Collection
    Set prevNodes = New Collection
End Sub


Public Function AddNextNode(ByRef node As c_Node) As c_Node
    AddUniqueCollectionItem nextNodes, node, node.number
Set AddNextNode = Me
End Function
Public Function AddPrevNode(ByRef node As c_Node) As c_Node
    AddUniqueCollectionItem prevNodes, node, node.number
Set AddPrevNode = Me
End Function

'Public Function AppendNextNode(ByRef node As c_Node) As c_Node
'    AddNextNode node
'    node.AddPrevNode Me
'Set AppendNextNode = Me
'End Function

Public Sub ClearCalculatedData()
    peoplesFlow = 0
    t_here = 0
    t_flow = 0
End Sub

Public Function RaisePeopleFlow(ByVal val As Integer) As c_Node
    peoplesFlow = peoplesFlow + val
Set RaisePeopleFlow = Me
End Function
Public Function RaiseTimeFlow(ByVal val As Single) As c_Node
    t_flow = t_flow + val
Set RaiseTimeFlow = Me
End Function



Public Property Get qw() As Single
'Произведение q*w
qw = q * width
End Property
Public Property Get qw_prevs() As Single
'Сумма q*w всех предыдущих узлов
Dim qw_prev As Single
Dim node As c_Node
    
    For Each node In prevNodes
        qw_prev = qw_prev + node.qw
    Next node
    
qw_prevs = qw_prev
End Property
