VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Rect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public pointBL As c_Point       'Coords of bottom left corner
'Public pointUR As c_Point       'Coords of upper top corner
'
'
'
'
'Public Sub SetData(ByRef point1 As c_Point, ByRef point2 As c_Point)
'    pointBL.SetEqual point1
'    pointUR.SetEqual point2
'End Sub
'
'Public Sub SetEqual(ByRef rect As c_Rect)
'    pointBL.SetEqual rect.pointBL
'    pointUR.SetEqual rect.pointUR
'End Sub


Public x As Double
Public y As Double
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
    width = shp.Cells("Width").Result(visInches)
    height = shp.Cells("Height").Result(visInches)
    angle = AngleToPage(shp)
End Sub

Public Function MakeNewRectShape() As Visio.Shape
Dim newShp As Visio.Shape
    
    Set newShp = Application.ActivePage.DrawRectangle(0, 0, 0, 0)
    
    newShp.Cells("PinX") = x
    newShp.Cells("PinY") = y
    newShp.Cells("Width") = width
    newShp.Cells("Height") = height * 1.05
    newShp.Cells("Angle") = angle
    Set MakeNewRectShape = newShp
End Function

Public Function MakeNewRectShapeInPos(xPos As Double, yPos As Double) As Visio.Shape
'Создаем новую фигуру в соответствии с указанной позицией
Dim newShp As Visio.Shape
    
    Set newShp = Application.ActivePage.DrawRectangle(0, 0, 0, 0)
    
    newShp.Cells("PinX") = xPos
    newShp.Cells("PinY") = yPos
    newShp.Cells("Width") = width
    newShp.Cells("Height") = height ' * 1.05
    newShp.Cells("Angle") = angle
    Set MakeNewRectShapeInPos = newShp
End Function

