VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public X As Double
Public Y As Double
Private nextPoint As c_Point        'Основная следующая точка
Private nextPoint2 As c_Point       'Резервная следующая точка
Private direction As Directions

Public Sub SetData(ByVal xVal As Double, ByVal yVal As Double)
    X = xVal
    Y = yVal
End Sub

Public Sub SetEqual(ByRef point As c_Point)
    X = point.X
    Y = point.Y
End Sub

Public Function isEqual(ByRef point As c_Point) As Boolean
    isEqual = X = point.X And Y = point.Y
End Function

Public Sub SetNextPoint(ByRef pnt As c_Point)
    If nextPoint Is Nothing Then
        Set nextPoint = pnt
    Else
        Set nextPoint2 = pnt
    End If
    CheckDir
End Sub
Public Function GetNextPoint() As c_Point
    If Not nextPoint2 Is Nothing Then
        Set GetNextPoint = nextPoint2
        Set nextPoint2 = Nothing
    Else
        Set GetNextPoint = nextPoint
        Set nextPoint = Nothing
    End If
End Function

Public Property Get Dir()
     Dir = direction
End Property

Private Sub CheckDir()
    If X = nextPoint.X And Y < nextPoint.Y Then
        direction = u
    End If
    If X = nextPoint.X And Y > nextPoint.Y Then
        direction = d
    End If
    If Y = nextPoint.Y And X < nextPoint.X Then
        direction = r
    End If
    If Y = nextPoint.Y And X > nextPoint.X Then
        direction = l
    End If
End Sub
