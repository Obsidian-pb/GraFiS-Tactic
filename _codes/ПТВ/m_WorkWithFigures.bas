Attribute VB_Name = "m_WorkWithFigures"
'----------------------------В модуле хранятся процедуры работы с формами набора--------------------------------------
Option Explicit

Public Sub PS_GlueToShape(ShpObj As Visio.Shape)
'Процедура привязывает инициировавшую фигуру (лафетный ствол) к целевой фигуре, в случае если она является _
 фигурой АЦ, АНР, Ар или прочее
Dim OtherShape As Visio.Shape
Dim x As Double, y As Double
Dim vS_ShapeName As String

    On Error GoTo EX

'---Определяем координаты активной фигуры
    x = ShpObj.Cells("PinX").Result(visInches)
    y = ShpObj.Cells("Piny").Result(visInches)

'---Проверяем налOtherShapeичие фигуры на месте перемещения лафетного ствола
    '---Перебираем все фигуры на странице
    For Each OtherShape In Application.ActivePage.Shapes
'    '---Если фигура является группой перебираем и все входящие в нее фигуры тоже
'        PS_GlueToShape OtherShape
    '---Проверяем, является ли эта фигура фигурой АЦ, АНР, Ар или прочее
    If GetTypeShape(OtherShape) > 0 And GetTypeShape(OtherShape) <> 24 And GetTypeShape(OtherShape) <> 30 And GetTypeShape(OtherShape) <> 31 Then
        '---Переводим координаты к координатам фигуры
'        OtherShape.XYFromPage x, y, x, y
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---Приклеиваем фигуру (прописываем формулы)
            On Error Resume Next
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." _
                & OtherShape.ID & "!Width*0.55)*SIN(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." _
                & OtherShape.ID & "!Width*0.55)*COS(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("Angle").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Angle+IF(Sheet." _
                & OtherShape.ID & "!User.DownOrient=1,-20 deg,20 deg))"
            ShpObj.Cells("Width").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Width*0.575)"
            ShpObj.Cells("Height").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Height*0.583)"
            ShpObj.Cells("Prop.Unit").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Unit)"
            ShpObj.Cells("Prop.SetTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.ArrivalTime"
            ShpObj.Cells("User.DownOrient").FormulaU = "Sheet." & OtherShape.ID & "!User.DownOrient"
            ShpObj.Cells("User.ShapeFromID").FormulaU = OtherShape.ID
            
            OtherShape.Cells("User.GFS_OutLafet").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.PodOut),0," & _
                "Sheet." & ShpObj.ID & "!User.PodOut)"
            OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.Head),0," & _
                "Sheet." & ShpObj.ID & "!User.Head)"
            OtherShape.BringToFront
        End If
    End If
    'для поезда
    If GetTypeShape(OtherShape) = 24 Then
        '---Переводим координаты к координатам фигуры
'        OtherShape.XYFromPage x, y, x, y
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---Приклеиваем фигуру (прописываем формулы)
            On Error Resume Next
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." _
                & OtherShape.ID & "!Width*0.55)*SIN(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." _
                & OtherShape.ID & "!Width*0.55)*COS(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("Angle").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Angle+IF(Sheet." _
                & OtherShape.ID & "!User.DownOrient=1,-20 deg,20 deg))"
            ShpObj.Cells("Width").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Width*0.6)"
            ShpObj.Cells("Height").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Height*0.4)"
            ShpObj.Cells("Prop.Unit").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Unit)"
            ShpObj.Cells("Prop.SetTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.ArrivalTime"
            ShpObj.Cells("User.DownOrient").FormulaU = "Sheet." & OtherShape.ID & "!User.DownOrient"
            ShpObj.Cells("User.ShapeFromID").FormulaU = OtherShape.ID
            
            OtherShape.Cells("User.GFS_OutLafet").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.PodOut),0," & _
                "Sheet." & ShpObj.ID & "!User.PodOut)"
            OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.Head),0," & _
                "Sheet." & ShpObj.ID & "!User.Head)"
            OtherShape.BringToFront
        End If
    End If
    
        'для судов
    If GetTypeShape(OtherShape) = 30 Or GetTypeShape(OtherShape) = 31 Then
        '---Переводим координаты к координатам фигуры
'        OtherShape.XYFromPage x, y, x, y
        If OtherShape.HitTest(x, y, 0.01) > 1 Then
            '---Приклеиваем фигуру (прописываем формулы)
            On Error Resume Next
            ShpObj.Cells("PinX").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinX-(Sheet." _
                & OtherShape.ID & "!Width*0.7)*SIN(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("PinY").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!PinY+(Sheet." _
                & OtherShape.ID & "!Width*0.7)*COS(Sheet." _
                & OtherShape.ID & "!Angle+90 deg)*IF(Sheet." & OtherShape.ID & "!User.DownOrient=1,-1,1))"
            ShpObj.Cells("Angle").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Angle+IF(Sheet." _
                & OtherShape.ID & "!User.DownOrient=1,-20 deg,20 deg))"
            ShpObj.Cells("Width").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Width*0.6)"
            ShpObj.Cells("Height").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Height*0.4)"
            ShpObj.Cells("Prop.Unit").FormulaU = "GUARD(Sheet." & OtherShape.ID & "!Prop.Unit)"
            ShpObj.Cells("Prop.SetTime").FormulaU = "Sheet." & OtherShape.ID & "!Prop.ArrivalTime"
            ShpObj.Cells("User.DownOrient").FormulaU = "Sheet." & OtherShape.ID & "!User.DownOrient"
            ShpObj.Cells("User.ShapeFromID").FormulaU = OtherShape.ID
            
            OtherShape.Cells("User.GFS_OutLafet").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.PodOut),0," & _
                "Sheet." & ShpObj.ID & "!User.PodOut)"
            OtherShape.Cells("User.GFS_OutLafet.Prompt").FormulaU = "IF(ISERR(Sheet." & ShpObj.ID & "!User.Head),0," & _
                "Sheet." & ShpObj.ID & "!User.Head)"
            OtherShape.BringToFront
        End If
    End If
Next OtherShape

Set OtherShape = Nothing
Exit Sub
EX:
    Set OtherShape = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
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
        vi_TempIndex = 11 Or vi_TempIndex = 20 Or vi_TempIndex = 24 Or vi_TempIndex = 30 Or _
        vi_TempIndex = 31 Then
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
'Процедура маняет направление потока в водосборнике
Public Sub DirectVS(shp As Visio.Shape, AsRazv As Boolean)
  If AsRazv Then ' Сборник
     shp.Cells("Connections.GFS_In").RowNameU = "GFS_Out"
     shp.Cells("Connections.GFS_Out1").RowNameU = "GFS_In1"
     shp.Cells("Connections.GFS_Out2").RowNameU = "GFS_In2"
     shp.Cells("Scratch.C1").FormulaU = 0
     shp.Cells("Scratch.D1").FormulaU = 0
     shp.Cells("Scratch.C2").FormulaU = "Scratch.C1+Prop.HeadLost"
     shp.Cells("Scratch.D2").FormulaU = "Scratch.D1/(Scratch.A3+Scratch.A2)"
     shp.Cells("Scratch.C3").FormulaU = "Scratch.C1+Prop.HeadLost"
     shp.Cells("Scratch.D3").FormulaU = "Scratch.D1/(Scratch.A3+Scratch.A2)"
  Else           ' Разветвитель
     shp.Cells("Connections.GFS_Out").RowNameU = "GFS_In"
     shp.Cells("Connections.GFS_In1").RowNameU = "GFS_Out1"
     shp.Cells("Connections.GFS_In2").RowNameU = "GFS_Out2"
     shp.Cells("Scratch.C1").FormulaU = "MAX(Scratch.C2,Scratch.C3)+Prop.HeadLost"
     shp.Cells("Scratch.D1").FormulaU = "Scratch.D2+Scratch.D3"
     shp.Cells("Scratch.C2").FormulaU = 0
     shp.Cells("Scratch.D2").FormulaU = 0
     shp.Cells("Scratch.C3").FormulaU = 0
     shp.Cells("Scratch.D3").FormulaU = 0
  End If
  shp.Cells("User.UseAsRazv").FormulaU = AsRazv
  
End Sub

Public Sub SearchWSShape(shp As Visio.Shape, x As Double, y As Double)
'Процедура ищет фигуры водоисточников и проверяет не оказался ли конец всасывающей линии в них. Если оказался, то для ячейки "User.WSShapeID" устанавливает ее ID
Dim curShp As Visio.Shape
Dim corX As Double
Dim corY As Double
        
    'Преобразуем локальные координаты фигуры в коорджинаты страницы
    shp.XYToPage corX, corY, x, y
    
    'перебираем все страницы на листе и проверяем, являются ли они фигурой Водоема
    For Each curShp In Application.ActivePage.Shapes
        'Если пожарный водоем
        If IsGFSShapeWithIP(curShp, 51, True) Then
            'Проверяем находится ли указанная точка внутри фигуры водоема
            If curShp.HitTest(x, y, 0) > 0 Then
                shp.Cells("User.WSShapeID").Formula = curShp.ID
                Exit Sub
            End If
        End If
        'Если открытый водоисточник
        If IsGFSShapeWithIP(curShp, 53, True) Then
            'Проверяем находится ли указанная точка внутри фигуры открытого водоисточника
            If curShp.HitTest(x, y, 0) > 0 Then
                shp.Cells("User.WSShapeID").Formula = curShp.ID
                Exit Sub
            End If
        End If
    Next curShp
    
'Если ничего не было найдено, указываем, что всасывающая линия ни откуда не забирает воду
shp.Cells("User.WSShapeID").Formula = 0
End Sub







