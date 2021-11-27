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
        ShpObj.Cells("Prop.ApearanceTime").FormulaU = "Sheet." & ToShape & "!Prop.LineTime" & ""
        ShpObj.Cells("Prop.Pressure").FormulaU = "ROUND(Sheet." & ToShape & "!Prop.HeadInHose*Prop.Koeff,2)" & ""
        
        RotateAtHoseLine ShpObj, ShpObj.Connects.Item(1).ToSheet
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.HoseDiameter").FormulaU = 0
        ShpObj.Cells("User.HoseNumber").FormulaU = 0
        ShpObj.Cells("User.WaterExpence").FormulaU = 0
        ShpObj.Cells("User.Resistance").FormulaU = 0
        ShpObj.Cells("User.LineLenight").FormulaU = 0
        ShpObj.Cells("Prop.ApearanceTime").FormulaU = 0
        ShpObj.Cells("Prop.Pressure").FormulaU = 0
    End If
    
End Sub

Public Sub ConnSoft(ShpObj As Visio.Shape)
'Процедура приклеивания подписи положения рукава
Dim ToShape As Integer

'---Предотвращаем появление сообщения об ошибке
On Error Resume Next

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("Prop.ApearanceTime").FormulaU = "Sheet." & ToShape & "!Prop.LineTime" & ""
        
        RotateAtHoseLine ShpObj, ShpObj.Connects.Item(1).ToSheet
    Else
        ShpObj.Cells("Prop.ApearanceTime").FormulaU = 0
    End If
    
End Sub

'-----------------------------------------Поворот фигуры на рукавной линии----------------------------------------------
Private Sub RotateAtHoseLine(ByRef lblShp As Visio.Shape, ByRef hoseLineShp As Visio.Shape)
'Основная процедура поворота фигуры на рукавной линии
Dim newCenter As c_Vector
Dim vector1 As c_Vector
Dim cll As Visio.Cell
Dim x As Double
Dim y As Double
    
    On Error GoTo EX
    
'    Application.EventsEnabled = False
    
    '1 находим точку на рукавной линии
    Set newCenter = New c_Vector
    newCenter.x = CellVal(lblShp, "BeginX")
    newCenter.y = CellVal(lblShp, "BeginY")
    
    '2 ищем ворую точку на рукавной линии
    Set vector1 = GetPointOnLineShape(newCenter, hoseLineShp, 0.1, 0)     ' lblShp.Cells("Width").Result(visInches) / 2, 0)
   
    
    '3 Переводим координаты страницы в координаты фигуры линии
    hoseLineShp.XYFromPage vector1.x, vector1.y, x, y
    
    '4 перемещаем второй коннектор в полученную точку
    Set cll = lblShp.Connects.Item(2).ToCell
    cll.FormulaU = "Width*" & Replace(CStr(x / hoseLineShp.Cells("Width")), ",", ".")
    Set cll = hoseLineShp.CellsSRC(visSectionConnectionPts, cll.Row, 1)
    cll.FormulaU = "Height*" & Replace(CStr(y / hoseLineShp.Cells("Height")), ",", ".")

EX:
'    Application.EventsEnabled = True
End Sub

