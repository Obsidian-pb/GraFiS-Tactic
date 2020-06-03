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
Private nextPoint As c_Point
Private direction As Directions

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

Public Sub SetNextPoint(ByRef pnt As c_Point)
    Set nextPoint = pnt
    CheckDir
End Sub
Public Function GetNextPoint() As c_Point
    Set GetNextPoint = nextPoint
End Function

Public Property Get Dir()
     Dir = direction
End Property

Private Sub CheckDir()
    If x = nextPoint.x And y < nextPoint.y Then
        direction = u
    End If
    If x = nextPoint.x And y > nextPoint.y Then
        direction = d
    End If
    If y = nextPoint.y And x < nextPoint.x Then
        direction = r
    End If
    If y = nextPoint.y And x > nextPoint.x Then
        direction = l
    End If
End Sub
