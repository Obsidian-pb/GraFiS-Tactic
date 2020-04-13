Attribute VB_Name = "m_Connects"
Option Explicit
'---------------------------------Модуль для хранения процедур линковки данных-----------------------------

Public Sub Conn(ShpObj As Visio.Shape)
'Процедура привязки отображаемого значения в подписи к фигуре к котрой она приклеена
Dim ToShape As Long

'---Предотвращаем появление сообщения об ошибке
On Error Resume Next

'---Если подпись ни к чему не приклеена, процедура заканчивается
    If ShpObj.Connects.Count > 0 Then
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = "Sheet." & ToShape & "!Prop.Street" & ""
    Else
        ToShape = ShpObj.Connects.Item(1).ToSheet.ID
        ShpObj.Cells("User.Label").FormulaU = 0
    End If

'---Перемещаем фигуру поверх прочих (На передний план)
    ShpObj.BringToFront
End Sub


'Процедура задает цвет линии расстояния
Public Sub DistBuild(ShpObj As Visio.Shape)
Dim FBeg As String, FEnd As String
Dim ShpSubj As Visio.Shape

    FBeg = ShpObj.Cells("BegTrigger").FormulaU
    FEnd = ShpObj.Cells("EndTrigger").FormulaU

'голубой
    If InStr(1, FBeg, "ПГ") <> 0 Or InStr(1, FEnd, "ПГ") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "Пирс") <> 0 Or InStr(1, FEnd, "Пирс") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "Колодец") <> 0 Or InStr(1, FEnd, "Колодец") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "Башня") <> 0 Or InStr(1, FEnd, "Башня") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "ПВ") <> 0 Or InStr(1, FEnd, "ПВ") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "водоисточник") <> 0 Or InStr(1, FEnd, "водоисточник") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If
    If InStr(1, FBeg, "Емкость") <> 0 Or InStr(1, FEnd, "Емкость") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
        GoTo SubjDistance
    End If

'фиолетовый
    If InStr(1, FBeg, "Здание") <> 0 And InStr(1, FEnd, "Здание") <> 0 Then
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD(RGB(112, 48, 160))"
        GoTo SubjDistance
    End If
    
Exit Sub
SubjDistance:
    ''---Передаем субъекту значение длины для маркирования фигур с указанным расстоянием
    Set ShpSubj = ActivePage.Shapes(Replace(Replace(FBeg, "_XFTRIGGER(", ""), "!EventXFMod)", ""))
    If ShpSubj.CellExists("User.Distance", 0) = False Then ShpSubj.AddNamedRow visSectionUser, "Distance", 0
    ShpSubj.Cells("User.Distance").FormulaU = "Sheet." & ShpObj.ID & "!Width"
    
    Set ShpSubj = ActivePage.Shapes(Replace(Replace(FEnd, "_XFTRIGGER(", ""), "!EventXFMod)", ""))
    If ShpSubj.CellExists("User.Distance", 0) = False Then ShpSubj.AddNamedRow visSectionUser, "Distance", 0
    ShpSubj.Cells("User.Distance").FormulaU = "Sheet." & ShpObj.ID & "!Width"
  
End Sub



Function InsertDistance(ShpObj As Visio.Shape, Optional Contex As Integer = 0)
'Процедура добавления strelki rasstoiania от соседних зданий
'---Объявляем переменные
Dim shpTarget As Visio.Shape
Dim shpConnection As Visio.Shape, vsO_Shape As Visio.Shape
Dim mstrConnection As Visio.Master, mstrSrelka As Visio.Master
Dim vsoCell1 As Visio.cell, vsoCell2 As Visio.cell
Dim CellFormula As String
Dim vsi_ShapeIndex As Integer
Dim lmax As Integer
Dim inppw As Boolean

vsi_ShapeIndex = 0

'    On Error GoTo EX
    InputDistanceForm.Show
    If InputDistanceForm.Flag = False Then Exit Function  'Если был нажат кэнсел - выходим не обновляя
    lmax = InputDistanceForm.lmax
    inppw = InputDistanceForm.inppw
    
    '---Перебираем все фигуры и находим здания
    For Each shpTarget In Application.ActivePage.Shapes
        If shpTarget.CellExists("User.IndexPers", 0) = True And shpTarget.CellExists("User.Version", 0) = True Then 'Является ли фигура фигурой ГраФиС
           'Проверяем не указано ли расстояние до фигуры уже
            If shpTarget.CellExists("User.Distance", 0) = False Then shpTarget.AddNamedRow visSectionUser, "Distance", 0
               If InStr(1, shpTarget.Cells("User.Distance").FormulaU, "!Width") = 0 Then
                vsi_ShapeIndex = shpTarget.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
                If vsi_ShapeIndex = 135 Then
                '---Вбрасываем коннектор и соединяем фигуру нашего здания и соседнего
                    Set mstrConnection = ThisDocument.Masters("Distance")
                    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
                  
                    Set vsoCell1 = shpConnection.CellsU("EndX")
                    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = shpConnection.CellsU("BeginX")
                    Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                ''---Проверяем длину линии и если она не подходит, удаляем
                    If shpConnection.Cells("Width").ResultIU = 0 Or shpConnection.Cells("Width").Result(visMeters) > lmax Then shpConnection.Delete
                End If

                If inppw = True And (vsi_ShapeIndex = 50 Or vsi_ShapeIndex = 51 Or vsi_ShapeIndex = 53 _
                   Or vsi_ShapeIndex = 54 Or vsi_ShapeIndex = 55 Or vsi_ShapeIndex = 56 Or vsi_ShapeIndex = 240 Or vsi_ShapeIndex = 190) Then
                     '---Вбрасываем коннектор и соединяем фигуру нашего здания и водоисточника
                    Set mstrConnection = ThisDocument.Masters("Distance")
                    Set shpConnection = Application.ActiveWindow.Page.Drop(mstrConnection, 2, 2)
                  
                    Set vsoCell1 = shpConnection.CellsU("EndX")
                    Set vsoCell2 = ShpObj.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    Set vsoCell1 = shpConnection.CellsU("BeginX")
                    Set vsoCell2 = shpTarget.CellsSRC(1, 1, 0)
                        vsoCell1.GlueTo vsoCell2
                    shpConnection.Cells("Prop.ArrowStyle").FormulaU = "INDEX(2,Prop.ArrowStyle.Format)"
                End If
               End If
        End If
     Next
     
'     If Contex = 0 And vsi_ShapeIndex = 0 Then Exit Function
    
      
'---Ставим фокус
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select ShpObj, visSelect

Exit Function
EX:
    SaveLog Err, "InsertDistance"
End Function




