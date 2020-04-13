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

    FBeg = ShpObj.Cells("BegTrigger").FormulaU
    FEnd = ShpObj.Cells("EndTrigger").FormulaU

'голубой
    If InStr(1, FBeg, "ПГ") <> 0 Or InStr(1, FEnd, "ПГ") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Пирс") <> 0 Or InStr(1, FEnd, "Пирс") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Колодец") <> 0 Or InStr(1, FEnd, "Колодец") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Башня") <> 0 Or InStr(1, FEnd, "Башня") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "ПВ") <> 0 Or InStr(1, FEnd, "ПВ") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "водоисточник") <> 0 Or InStr(1, FEnd, "водоисточник") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
    If InStr(1, FBeg, "Емкость") <> 0 Or InStr(1, FEnd, "Емкость") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD (RGB(0, 176, 240))"
'фиолетовый
    If InStr(1, FBeg, "Здание") <> 0 And InStr(1, FEnd, "Здание") <> 0 Then _
        ShpObj.Cells("LineColor").FormulaU = "THEMEGUARD(RGB(112, 48, 160))"
  
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
'            If shpTarget.Cells("User.Version") >= CP_GrafisVersion Then  'Проверяем версию фигуры
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

                If inppw = True And vsi_ShapeIndex = 50 Then
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
'            End If
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




