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

Private Const radToGrad = 57.2958





'===In work===
'Public Property Get Angle() As Variant
''Возвращает угол вектора в градусах
'    If x = 0 Then
'        If y > 0 Then
'            Angle = 90
'        ElseIf y = 0 Then
'            Angle = 0
'        ElseIf y < 0 Then
'            Angle = 270
'        End If
'        Exit Function
'    End If
'    If y = 0 Then
'        If x > 0 Then
'            Angle = 0
'        ElseIf x = 0 Then
'            Angle = 0
'        ElseIf x < 0 Then
'            Angle = 180
'        End If
'        Exit Function
'    End If
'
'    Angle = Atn(y / x) * radToGrad
'End Property
'
'Public Property Let Angle(ByVal ang As Double)
''Выисляет и устанавливает координаты x и y для нового угла
'    x = Cos(ang)
'    y = Sin(ang)
'End Property
