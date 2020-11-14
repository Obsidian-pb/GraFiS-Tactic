VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_NRSNodeObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------Демон-наблюдатель для узла НРС------------------------
'Private WithEvents node_ As c_NRSNode
Private node_ As c_NRSNode
Attribute node_.VB_VarHelpID = -1
Private history() As String                 'Содержание истории
Private historyPropID() As Integer          'Коды данных истории
'Private prop_ As NRSProp

Const delimiter = ","
Const decimalSign = "."



Public Function Activate(ByRef node As c_NRSNode) As c_NRSNodeObserver
    Set node_ = node
Set Activate = Me
End Function

Public Function AddProp(ByVal a_prop As NRSProp) As c_NRSNodeObserver
Dim newPropIndex As Integer
'    Stop
    'Здесь нужно специальную процедуру
    newPropIndex = GetNextHistoryIndex()   ' UBound(history, 1) + 1
'    ReDim history(newPropIndex, 1)
    historyPropID(newPropIndex) = a_prop
    
Set AddProp = Me
End Function


'Private Sub Class_Initialize()
'    ReDim history(0, 1)
'End Sub

'Private Sub node__PropChanged(ByVal val As Single, ByVal changedProp As NRSProp)
'Dim i As Integer
''    Debug.Assert Not changedProp = nrsProphOut
''    If changedProp = nrsProphOut Then
''        Debug.Assert Not changedProp = nrsProphOut
''    End If
'    For i = 0 To UBound(history)
'        If historyPropID(i) = changedProp Then
'            history(i) = history(i) & Replace(CStr(val), ",", ".") & delimiter
'            Exit Sub
'        End If
'    Next i
'
''    If changedProp = prop_ Then
''        history = history & CStr(val) & delimiter
''    End If
'
'End Sub

Public Sub HistoryExpand()
Dim i As Integer
Dim val As Single
Dim valStr As String
   
    For i = 0 To UBound(historyPropID)
        Select Case historyPropID(i)
            Case Is = NRSProp.nrsPropH
                val = node_.H
            Case Is = NRSProp.nrsProphIn
                val = node_.hIn
            Case Is = NRSProp.nrsProphLost
                val = node_.hLost
            Case Is = NRSProp.nrsProphOut
                val = node_.hOut
            Case Is = NRSProp.nrsPropP
                val = node_.P
            Case Is = NRSProp.nrsPropQ
                val = node_.Q
            Case Is = NRSProp.nrsPropS
                val = node_.S
            Case Is = NRSProp.nrsPropZ
                val = node_.Z
        End Select
        
        valStr = Replace(CStr(val), ",", ".")
        history(i) = history(i) & valStr & delimiter
    Next i

End Sub

Public Property Get NRSNode() As Variant
    Set NRSNode = node_
End Property
Public Property Set NRSNode(ByVal nd As c_NRSNode)
    Set node_ = nd
End Property

Public Property Get ID() As String
    ID = node_.ID
End Property

Public Function GetHistory(ByVal a_prop As NRSProp) As String
Dim i As Integer
Dim hist As String
    
    For i = 0 To UBound(history)
        If historyPropID(i) = a_prop Then
            hist = history(i)
            GetHistory = Left(hist, Len(hist) - 1)
            Exit Function
        End If
    Next i

    
End Function



Private Function GetNextHistoryIndex() As Integer
Dim newPropIndex As Integer
    On Error GoTo ex
    newPropIndex = UBound(history) + 1
    ReDim Preserve history(newPropIndex)
    ReDim Preserve historyPropID(newPropIndex)
    GetNextHistoryIndex = newPropIndex
Exit Function
ex:
    ReDim history(0)
    ReDim historyPropID(0)
    GetNextHistoryIndex = 0
End Function
