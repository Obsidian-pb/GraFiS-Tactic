VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------Модель насосно-рукавной системы-----------------------
Private NRSNodes_ As Collection                         'Коллекция узлов НРС
Private starters_ As Collection                         'Коллекция питателей НРС
Private enders_ As Collection                           'Коллекция потребителей НРС

Public correct As Boolean                              'Флаг корректности модели

Private Sub Class_Initialize()
    Set NRSNodes_ = New Collection
    Set starters_ = New Collection
    Set enders_ = New Collection
    correct = True
End Sub
Private Sub Class_Terminate()
    Set NRSNodes_ = Nothing
    Set starters_ = Nothing
    Set enders_ = Nothing
End Sub





Public Property Get NRSNodes() As Collection
    Set NRSNodes = NRSNodes_
End Property
Public Function AddNRSNode(ByRef node As c_NRSNode) As c_NRSModel
    AddUniqueCollectionItem NRSNodes_, node
    If node.isStarter Then
        AddUniqueCollectionItem starters_, node
    ElseIf node.isEnder Then
        AddUniqueCollectionItem enders_, node
    End If

Set AddNRSNode = Me
End Function
Public Function DeleteNRSNode(ByRef node As c_NRSNode) As c_NRSModel
    RemoveFromCollection NRSNodes_, node

Set DeleteNRSNode = Me
End Function
Public Function Clear() As c_NRSModel
    Set NRSNodes_ = New Collection

Set Clear = Me
End Function

Public Property Get StarterNodes() As Collection
    Set StarterNodes = starters_
End Property

Public Property Get EnderNodes() As Collection
    Set EnderNodes = enders_
End Property

Public Function GetNodesCount() As Long
    GetNodesCount = NRSNodes_.Count
End Function

Public Function InModel(ByRef shp As Visio.Shape) As Boolean
    InModel = IsInCollection(NRSNodes_, shp)
End Function

Public Function GetNRSNodeByKey(ByVal NRSNodeID As String) As c_NRSNode
    On Error GoTo ex
    Set GetNRSNodeByKey = NRSNodes_.item(NRSNodeID)
    
Exit Function
ex:
    Set GetNRSNodeByKey = Nothing
'    Debug.Print "Such node is not existed in model"
End Function
Public Function GetNRSNodeByShape(ByRef shp As Visio.Shape) As c_NRSNode
    On Error GoTo ex
    If InModel(shp) Then
        Set GetNRSNodeByShape = GetNRSNodeByKey(CStr(shp.ID))
    End If
    
Exit Function
ex:
    Set GetNRSNodeByShape = Nothing
End Function

Public Function GetFirstNRSNode() As c_NRSNode
    On Error GoTo ex
    Set GetFirstNRSNode = NRSNodes_(1)
Exit Function
ex:
    Set GetFirstNRSNode = Nothing
End Function

Public Property Get TotalQ() As Single
Dim node As c_NRSNode
Dim sum As Single
    
    For Each node In EnderNodes
        sum = sum + node.Q
    Next node
    
TotalQ = sum
End Property
