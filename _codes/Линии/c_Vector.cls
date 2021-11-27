VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public x As Double
Public y As Double
Public segmentNumber As Byte                'Ќомер сегмента круга в котором вектор был получен

Private Const radToGrad = 57.2958


Private Sub Class_Initialize()
    segmentNumber = 0
End Sub



Public Property Get Angle() As Double
'¬озвращает угол вектора в градусах
    If x = 0 Then
        If y > 0 Then
            Angle = 90
        ElseIf y = 0 Then
            Angle = 0
        ElseIf y < 0 Then
            Angle = 270
        End If
        Exit Property
    End If
    If y = 0 Then
        If x > 0 Then
            Angle = 0
        ElseIf x = 0 Then
            Angle = 0
        ElseIf x < 0 Then
            Angle = 180
        End If
        Exit Property
    End If

    Angle = Atn(y / x) * radToGrad
End Property

Public Property Let Angle(ByVal ang As Double)
'¬ыисл€ет и устанавливает координаты x и y дл€ нового угла
    x = Cos(ang)
    y = Sin(ang)
End Property

Public Function IsSame(ByRef otherVector As c_Vector, Optional ByVal tolerance As Double = 0) As Boolean
'¬озвращает True, если данный вектор совпадает с вектором otherVector с учтом чувствительности tolerance
    IsSame = x >= otherVector.x - tolerance And _
             x <= otherVector.x + tolerance And _
             y >= otherVector.y - tolerance And _
             y <= otherVector.y + tolerance
End Function


