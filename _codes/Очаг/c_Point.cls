VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double

Public Sub SetData(ByVal xVal As Double, ByVal yVal As Double)
    x = xVal
    y = yVal
End Sub

Public Sub SetEqual(ByRef point As c_Point)
    x = point.x
    y = point.y
End Sub

Public Function isEqual(ByRef point As c_Point) As Boolean
    isEqual = x = point.x And y = point.y
End Function

