VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Rect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double
'Public locPinX As Double
'Public locPinY As Double
'Public locPinX As String
'Public locPinY As String
Public width As Double
Public height As Double
Public angle As Double



Private Sub Class_Initialize()
    x = 0
    y = 0
    locPinX = 0
    locPinY = 0
    width = 0
    height = 0
    angle = 0
End Sub

Public Function GetRectShape(shp As Visio.Shape) As Visio.Shape
    ImportFromShape shp
    Set GetRectShape = MakeNewRectShape
End Function

Public Sub ImportFromShape(shp As Visio.Shape)
    shp.XYToPage shp.Cells("LocPinX").Result(visInches), shp.Cells("LocPinY").Result(visInches), x, y
'    locPinX = shp.Cells("LocPinX").Result(visInches)
'    locPinY = shp.Cells("LocPinY").Result(visInches)
'    locPinX = shp.Cells("LocPinX").Formula
'    locPinY = shp.Cells("LocPinY").Formula
    width = shp.Cells("Width").Result(visInches)
    height = shp.Cells("Height").Result(visInches)
    angle = AngleToPage(shp)
End Sub

Public Function MakeNewRectShape() As Visio.Shape
Dim newShp As Visio.Shape
    
    Set newShp = Application.ActivePage.DrawRectangle(0, 0, 0, 0)
    
    newShp.Cells("PinX") = x
    newShp.Cells("PinY") = y
'    newShp.Cells("locPinX") = locPinX
'    newShp.Cells("locPinY") = locPinY
'    newShp.Cells("locPinX").Formula = locPinX
'    newShp.Cells("locPinY").Formula = locPinY
    newShp.Cells("Width") = width
    newShp.Cells("Height") = height * 1.05
    newShp.Cells("Angle") = angle
    Set MakeNewRectShape = newShp
End Function

