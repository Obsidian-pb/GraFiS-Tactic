VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public startTime As Date


Private Sub Class_Initialize()
    startTime = Now()
End Sub

Public Sub PrintElapsedTime()
    Debug.Print "elapsed " & DateDiff("s", startTime, Now()) & " seconds"
    startTime = Now()
End Sub

Public Function GetElapsedTime() As Integer
    GetElapsedTime = DateDiff("s", startTime, Now())
End Function
