Attribute VB_Name = "Connects"

Public Sub Conn(ShpObj As Visio.Shape)
'Процедура привязки отображаемого значения в подписи к фигуре к котрой она приклеена
Dim ToShape As Integer

'---Предотвращаем появление сообщения об ошибке
On Error Resume Next

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.HoseDiameter").FormulaU = "Sheet." & ToShape & "!Prop.HoseDiameter" & ""
        ShpObj.Cells("User.HoseNumber").FormulaU = "Sheet." & ToShape & "!User.HosesNeed" & ""
        ShpObj.Cells("User.WaterExpence").FormulaU = "Sheet." & ToShape & "!Prop.Flow.Value" & ""
        ShpObj.Cells("User.Resistance").FormulaU = "Sheet." & ToShape & "!Prop.HoseResistance.Value" & ""
        ShpObj.Cells("User.LineLenight").FormulaU = "Sheet." & ToShape & "!User.TotalLenight" & ""
        ShpObj.Cells("Prop.Pressure").FormulaU = "ROUND(Sheet." & ToShape & "!Prop.HeadInHose*Prop.Koeff,2)" & ""
        
        RotateAtHoseLine ShpObj, ShpObj.Connects.Item(1).ToSheet
    Else
'        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
'        ShpObj.Cells("User.HoseDiameter").FormulaU = 0
'        ShpObj.Cells("User.HoseNumber").FormulaU = 0
'        ShpObj.Cells("User.WaterExpence").FormulaU = 0
'        ShpObj.Cells("User.Resistance").FormulaU = 0
'        ShpObj.Cells("User.LineLenight").FormulaU = 0
'        ShpObj.Cells("Prop.Pressure").FormulaU = 0
    End If
    
End Sub

'-----------------------------------------Поворот фигуры на рукавной линии----------------------------------------------
Private Sub RotateAtHoseLine(ByRef lblShp As Visio.Shape, ByRef hoseLineShp As Visio.Shape)
'Основная процедура поворота фигуры на рукавной линии
Dim newCenter As c_Vector
Dim vector1 As c_Vector
Dim vector2 As c_Vector
'Dim vector3 As c_Vector
    
'    On Error GoTo EX
    
'    Application.EventsEnabled = False
    
    '1 находим точку на рукавной линии
'    Set newCenter = FindNearestHoseLinePoint(curPoint, hoseLineShp)
'    If newCenter Is Nothing Then Exit Sub
    Set newCenter = New c_Vector
    newCenter.x = CellVal(lblShp, "BeginX")
    newCenter.y = CellVal(lblShp, "BeginY")
'    lblShp.Cells("EndX").Formula = newCenter.x + 0.01
'    lblShp.Cells("EndY").Formula = newCenter.y + 0.01
    
'    'Проверяем, не находится ли уже фигура в данной точке
'    If curPoint.IsSame(newCenter, 0.1) Then Exit Sub
    
'    '2 размещаем фигуру в ближайшей точке
'    lblShp.Cells("PinX").Formula = str(newCenter.x)
'    lblShp.Cells("PinY").Formula = str(newCenter.y)
    
    '3 ищем две точки на линии
    Set vector1 = GetPointOnLineShape(newCenter, hoseLineShp, 10, 0)    ' lblShp.Cells("Width").Result(visInches) / 2, 0)
'    Set vector2 = GetPointOnLineShape(curPoint, hoseLineShp, lblShp.Cells("Height").Result(visInches) / 2, vector1.segmentNumber + 1)
'    Set vector3 = NewVectorXY(vector1.x - vector2.x, vector1.y - vector2.y)
'    lblShp.Cells("Angle").Formula = str(vector3.Angle) & "deg"
    
'    Debug.Print lblShp.Connects.Item(2).ToCell.Name   '  .FromCell.Name
'    hoseLineShp.DeleteRow visSectionConnectionPts, lblShp.Connects.Item(1).ToCell.Row
    hoseLineShp.DeleteRow visSectionConnectionPts, lblShp.Connects.Item(2).ToCell.Row
    lblShp.Cells("EndX").Formula = vector1.x   ' "BeginX+" & vector1.x - newCenter.x
    lblShp.Cells("EndY").Formula = vector1.y   ' "BeginY+" & vector1.y - newCenter.y
    
'''
'''
''''    Set vector2 = New c_Vector
'''    Dim x As Double
'''    Dim y As Double
'''    Dim frml As String
'''
''''    hoseLineShp.AddRow visSectionConnectionPts, hoseLineShp.RowCount(visSectionConnectionPts) + 1, 0
'''    hoseLineShp.XYFromPage newCenter.x, newCenter.y, x, y
'''    frml = "PAR(PNT(Sheet." & hoseLineShp.ID & "!Width*" & Replace(CStr(x / hoseLineShp.Cells("Width")), ",", ".") & _
'''                                     ",Sheet." & hoseLineShp.ID & "!Height*" & Replace(CStr(y / hoseLineShp.Cells("Height")), ",", ".") & "))"
''''    frml = Replace(frml, ",", ".")
'''    lblShp.Cells("BeginX").FormulaU = frml
'''    lblShp.Cells("BeginY").FormulaU = frml
'''
'''    hoseLineShp.XYFromPage vector1.x, vector1.y, x, y
''''    frml = "PAR(PNT(Sheet." & hoseLineShp.ID & "!Width*" & x / hoseLineShp.Cells("Width") & _
''''                                     ",Sheet." & hoseLineShp.ID & "!Height*" & y / hoseLineShp.Cells("Height") & "))"
'''    frml = "PAR(PNT(Sheet." & hoseLineShp.ID & "!Width*" & Replace(CStr(x / hoseLineShp.Cells("Width")), ",", ".") & _
'''                                     ",Sheet." & hoseLineShp.ID & "!Height*" & Replace(CStr(y / hoseLineShp.Cells("Height")), ",", ".") & "))"
''''    frml = Replace(frml, ",", ".")
'''    lblShp.Cells("EndX").FormulaU = frml
'''    lblShp.Cells("EndY").FormulaU = frml
'''
'''
'''
''''    PAR(PNT(Sheet.34!Width*0.4,Sheet.34!Height*0.4))
'''
''''    hoseLineShp.CellsSRC(visSectionConnectionPts, hoseLineShp.RowCount(visSectionConnectionPts) - 1, 0).Formula = "Width*" & x / hoseLineShp.Cells("Width")
''''    hoseLineShp.CellsSRC(visSectionConnectionPts, hoseLineShp.RowCount(visSectionConnectionPts) - 1, 1).Formula = "Height*" & y / hoseLineShp.Cells("Height")
''''    lblShp.Cells("EndX").GlueTo hoseLineShp.CellsSRC(visSectionConnectionPts, hoseLineShp.RowCount(visSectionConnectionPts) - 1, 0)
'''
EX:
'    Application.EventsEnabled = True
End Sub


'Public Function FindNearestHoseLinePoint(ByRef curPoint As c_Vector, ByRef hoseLineShp As Visio.Shape) As c_Vector
''Находим ближайшую точку рукавной линии
'Dim distance As Double
'
'    '1 Находим расстояние до рукавной линии
'    distance = hoseLineShp.DistanceFromPoint(curPoint.x, curPoint.y, visSpatialIncludeDataGraphics)
'    '2 На найденном расстоянии находим точку
'    Set FindNearestHoseLinePoint = GetPointOnLineShape(curPoint, hoseLineShp, distance)
'
'End Function
