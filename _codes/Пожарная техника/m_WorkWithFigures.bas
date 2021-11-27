Attribute VB_Name = "m_WorkWithFigures"
'----------------------------В модуле хранятся процедуры работы с формами набора--------------------------------------
Option Explicit

Public Sub PS_GlueToShape(ShpObj As Visio.Shape)
'Процедура привязывает инициировавшую фигуру к целевой фигуре, в случае если она является _
 фигурой АЦ, АНР, Ар или прочее
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double
Dim vS_ShapeName As String
Dim TrueShape As Boolean

    On Error GoTo ex
TrueShape = False
'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---Проверяем налOtherShapeичие фигуры на месте перемещения лафетного ствола
    '---Перебираем все фигуры на странице
    For Each OtherShape In Application.ActivePage.Shapes
'    '---Если фигура является группой перебираем и все входящие в нее фигуры тоже
'        PS_GlueToShape OtherShape
    '---Проверяем, является ли эта фигура фигурой АЦ, АНР, Ар или прочее
    If GetTypeShape(OtherShape) > 0 Then
         
        '---Переводим координаты к координатам фигуры
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---Приклеиваем фигуру (прописываем формулы)
            On Error Resume Next
            TrueShape = True
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." & OtherShape.ID & "!Height+Sheet." & OtherShape.ID & "!Width)*0.25)"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." & OtherShape.ID & "!Height+Sheet." & OtherShape.ID & "!Width)*0.3)"
            ShpObj.Cells("Width").FormulaU = "GUARD((Sheet." & OtherShape.ID & "!Height+Sheet." & OtherShape.ID & "!Width)*0.35)"
            ShpObj.Cells("Height").FormulaU = "GUARD((Sheet." & OtherShape.ID & "!Height+Sheet." & OtherShape.ID & "!Width)*0.25)"
            ShpObj.Cells("LocPinX").FormulaU = "GUARD(Width*0.5)"
            ShpObj.Cells("LocPinY").FormulaU = "GUARD(Height*0.5)"
            ShpObj.Cells("Char.Size").FormulaU = "(Sheet." & OtherShape.ID & "!Height+Sheet." & OtherShape.ID & "!Width)*0.2/(ThePage!DrawingScale/ThePage!PageScale)"
'            ShpObj.Cells("Char.Color").FormulaU = "Sheet." & OtherShape.ID & "!User.ShapeColor1"
            ShpObj.Cells("User.NameModel").FormulaU = "Sheet." & OtherShape.ID & "!Prop.Model"
            ShpObj.Cells("User.MotherIndexPers").FormulaU = "Sheet." & OtherShape.ID & "!User.IndexPers"
            ShpObj.Cells("User.NameMother").FormulaU = """" & OtherShape.Name & """"
            ShpObj.Cells("Prop.NModel").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Model)"
            ShpObj.Cells("Prop.PersonnelHave").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.PersonnelHave*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses38").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose38*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses51").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose51*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses66").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose66*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses77").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose77*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses89").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose89*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses110").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose110*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses150").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose150*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses200").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose200*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses250").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose250*Prop.Quantity)"
            ShpObj.Cells("Prop.Hoses300").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Hose300*Prop.Quantity)"
            ShpObj.Cells("Prop.Powder").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Powder*Prop.Quantity)"
            ShpObj.Cells("Prop.Water").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Water*Prop.Quantity)"
            ShpObj.Cells("Prop.FoamX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Foam*Prop.Quantity)"
            ShpObj.Cells("Prop.Dest").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Dest)"
            ShpObj.Cells("Prop.RSCHS").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.RSCHS)"
            
            Exit For
        End If
    End If
Next OtherShape
If TrueShape = False Then ShpObj.Delete
Set OtherShape = Nothing
Exit Sub
ex:
    Set OtherShape = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PS_GlueToShape"
End Sub



Private Function GetTypeShape(ByRef aO_TergetShape As Visio.Shape) As Integer
'Функция возвращает индехс целевой фигуры в случае если она является фигурой приемлемой _
 для установки лафетного ствола, в противном случае возвращается 0
Dim vi_TempIndex

GetTypeShape = 0

If aO_TergetShape.CellExists("User.IndexPers", 0) = True And aO_TergetShape.CellExists("User.Version", 0) = True Then
    vi_TempIndex = aO_TergetShape.Cells("User.IndexPers")
    If vi_TempIndex = 1 Or vi_TempIndex = 2 Or vi_TempIndex = 9 Or vi_TempIndex = 10 Or _
        vi_TempIndex = 11 Or vi_TempIndex = 20 Or vi_TempIndex = 163 Or vi_TempIndex = 3 Or _
        vi_TempIndex = 4 Or vi_TempIndex = 161 Or vi_TempIndex = 162 Or vi_TempIndex = 5 Or _
        vi_TempIndex = 6 Or vi_TempIndex = 7 Or vi_TempIndex = 8 Or vi_TempIndex = 10 Or _
        vi_TempIndex = 11 Or vi_TempIndex = 12 Or vi_TempIndex = 13 Or vi_TempIndex = 14 Or _
        vi_TempIndex = 15 Or vi_TempIndex = 16 Or vi_TempIndex = 17 Or vi_TempIndex = 18 Or _
        vi_TempIndex = 19 Or vi_TempIndex = 160 Or vi_TempIndex = 73 Or vi_TempIndex = 74 Or _
        vi_TempIndex = 3000 Or vi_TempIndex = 3001 Or vi_TempIndex = 3002 Or vi_TempIndex = 24 Or _
        vi_TempIndex = 25 Or vi_TempIndex = 26 Or vi_TempIndex = 27 Or vi_TempIndex = 28 Or vi_TempIndex = 29 Or _
        vi_TempIndex = 30 Or vi_TempIndex = 31 Or vi_TempIndex = 32 Or vi_TempIndex = 33 _
    Then
        GetTypeShape = aO_TergetShape.Cells("User.IndexPers")
    End If
End If

End Function


Public Sub PS_ReleaseShape(ShpObj As Visio.Shape, ShapeID As Long)
'Процедура снимает закрепление лафетного ствола за автомобилем
Dim OtherShape As Visio.Shape

If ShapeID = 0 Then Exit Sub

Set OtherShape = Application.ActivePage.Shapes.ItemFromID(ShapeID)

On Error Resume Next

    ShpObj.Cells("PinX").FormulaForce = ShpObj.Cells("PinX").Result(visNumber)
    ShpObj.Cells("PinY").FormulaForce = ShpObj.Cells("PinY").Result(visNumber)
    ShpObj.Cells("Angle").FormulaForce = ShpObj.Cells("Angle").Result(visNumber)
    ShpObj.Cells("Width").FormulaForce = ShpObj.Cells("Width").Result(visNumber)
    ShpObj.Cells("Height").FormulaForce = ShpObj.Cells("Height").Result(visNumber)
    ShpObj.Cells("Prop.Unit").FormulaForceU = """" & ShpObj.Cells("Prop.Unit").ResultStr(visUnitsString) & """"
    ShpObj.Cells("Prop.SetTime").FormulaForce = ShpObj.Cells("Prop.SetTime").Result(visDate)
    ShpObj.Cells("User.DownOrient").FormulaU = "IF(Angle>-1 deg,1,0)"
    ShpObj.Cells("User.ShapeFromID").FormulaU = 0
    
    OtherShape.Cells("User.GFS_OutLafet").FormulaU = 0
    OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = 0
    ShpObj.BringToFront

Set OtherShape = Nothing
End Sub
