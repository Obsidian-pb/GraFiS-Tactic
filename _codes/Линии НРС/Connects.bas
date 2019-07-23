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
        ShpObj.Cells("Char.Color").FormulaU = "Sheet." & ToShape & "!LineColor" & ""
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.HoseDiameter").FormulaU = 0
        ShpObj.Cells("User.HoseNumber").FormulaU = 0
        ShpObj.Cells("User.WaterExpence").FormulaU = 0
        ShpObj.Cells("User.Resistance").FormulaU = 0
        ShpObj.Cells("Char.Color").FormulaU = "Styles!Р_Подпись!Char.Color"
    End If

End Sub


Public Sub NormalizeSystem(ShpObj As Visio.Shape)
    NormalizeNRS
End Sub
Public Sub NormalizeNRS()
'Процедура нормализации насосно-рукавной системы
'Позволяет обработать ошибки вычислений
Dim shp As Visio.Shape
Dim frml As String


    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) = True Then
            If shp.Cells("User.IndexPers") = 100 Then
                CellFix shp.Cells("Scratch.C1")
'                shp.Cells("Scratch.C1").Formula = ""
'                shp.Cells("Scratch.C1").Formula = frml
            End If
            If shp.Cells("User.IndexPers") = 34 Then
                CellFix shp.Cells("Scratch.A1")
                CellFix shp.Cells("User.PodOut")
                CellFix shp.Cells("User.Head")
'                shp.Cells("Scratch.A1").Formula = ""
'                shp.Cells("Scratch.A1").Formula = frml
            End If
        End If
    Next shp

End Sub

Private Sub CellFix(ByRef cell As Visio.cell)
Dim frml As String
    
    frml = cell.Formula
    cell.FormulaForce = ""
    cell.FormulaForce = frml
End Sub
