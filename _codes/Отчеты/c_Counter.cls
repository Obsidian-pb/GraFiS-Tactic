VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public val As Single
Public ID As String



Private Sub Class_Initialize()
    val = 0
End Sub

Public Sub Activate(ByVal a_id As String)
    ID = a_id
End Sub



Public Sub Rize(v)
    val = val + v
End Sub

Public Sub Del()
    val = 0
End Sub

