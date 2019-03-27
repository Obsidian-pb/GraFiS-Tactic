VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Fire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------------------Класс для процедур построения площади пожара в условиях сложной геометрии-----------------------------------
Dim WithEvents vO_App As Visio.Application
Attribute vO_App.VB_VarHelpID = -1
Private pO_MainFireShape As Visio.Shape 'Фигура основной площади горения - к ней добавляются дополнительные построения
Private pO_MainShape As Visio.Shape
Private pO_MaskShape As Visio.Shape 'фигура маски - суммируется из теней стен и используется как общая тень стен для обрезки
Dim VI_ShapesCheckFlag As Integer 'Флаг включает/выключает учет вбрасываемых фигур
Dim vB_ShapeAlreadyDetectFlag As Boolean 'Флаг включается в случае, если фигура уже была определена как pO_MainShape
                                         'во время текущей сессии вброса (при фрагментировании
Dim VD_MinSquare As Double
Dim x As Double, y As Double
Dim VI_PointsCount As Integer 'Количество точек от которых строятся фигуры
'Dim VI_ShapesByPointsCount As Integer 'Количество построенных от точек фигур
Dim v_DoorsCount As Integer 'Текущее количество дверей охваченных огнем
Dim VD_ShapeArea As Double  'Площадь фигур

'--------------------------Внешние свойства класса------------------------------------------------------------------------------
Public PS_LineSpeedM As Single 'Указывает линейную скорость распространения огня в метрах
'--------------------------Свойства для работы с объетами дверей----------------------------------------------------------------
Public PB_CheckOpens As Boolean 'Указывает, осуществляется ли проверка открытия/закрытия дверей
Public PI_DoorsREI As Integer   'Указывает, предел огнестойкости дверей по умолчанию


'--------------------------Коллекции--------------------------------------------------------------------------------------------
Dim Col_Points(300, 2) As Double 'Массив координат критических точек
Const conI_TopPoint = 300
Dim arrDoorsIDArray() As Integer
Private col_FireShapes As Collection 'Коллекция вброшенных фигур вторичных очагов
Private col_WallsShapes As Collection 'Коллекция стен находящихся в потенциальной досягаемости пожара
Private col_PortalShapes As Collection 'Коллекция проходов находящихся в потенциальной досягаемости пожара
Private col_DoorsShapes As Collection 'Коллекция дверей находящихся в потенциальной досягаемости пожара
Private col_DoorsFireShapes As Collection 'Коллекция вторичных очагов образованных дверными проемами






'--------------------------Процедуры класса-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set vO_App = Visio.Application
    VI_ShapesCheckFlag = 0
    vB_ShapeAlreadyDetectFlag = False
End Sub

Private Sub Class_Terminate()
    On Error Resume Next 'Отключен обработчик!!!
    pO_MaskShape.Delete 'удаляем маску
    Set vO_App = Nothing
End Sub

Private Sub vO_App_ShapeAdded(ByVal Shape As IVShape)
'В случае, если включен учет добавляемых фигур - проверяет является ли вброшенная фигура затравочной
'т.е. отходящей непосредственно от очага

    On Error GoTo EX
    Select Case VI_ShapesCheckFlag
        Case Is = 0
            Exit Sub
        Case Is = 1
            If Shape.HitTest(x, y, 0.01) = 0 Then
                'Предотвращаем некорректную обрезку
    '            If Shape.AreaIU < VD_ShapeArea * 1.01 And Shape.AreaIU > VD_ShapeArea * 0.99 Then Exit Sub
                Shape.Delete
            Else
                If vB_ShapeAlreadyDetectFlag = False Then
                    Set pO_MainShape = Shape
                    Application.ActiveWindow.Select pO_MainShape, visSelect
                    vB_ShapeAlreadyDetectFlag = True
                Else
                    If Shape.AreaIU > pO_MainShape.AreaIU Then
                        VD_MinSquare = pO_MainShape.AreaIU
                        pO_MainShape.Delete
                        Set pO_MainShape = Shape
                        Application.ActiveWindow.Select pO_MainShape, visSelect
                    Else
                        VD_MinSquare = Shape.AreaIU
                        Shape.Delete
                    End If
                End If
            End If
        Case Is = 2
            'Предотвращаем некорректную обрезку
    '        If Shape.AreaIU < VD_ShapeArea * 1.01 And Shape.AreaIU > VD_ShapeArea * 0.99 Then Exit Sub
            'обрабатываем основное событие
            If Shape.HitTest(x, y, 0.01) = 0 Then
                Shape.Delete
            Else
                If vB_ShapeAlreadyDetectFlag = False Then
                    Set pO_MainShape = Shape
                    Application.ActiveWindow.Select pO_MainShape, visSelect
                    vB_ShapeAlreadyDetectFlag = True
                Else
                    If Shape.AreaIU > pO_MainShape.AreaIU Then
                        VD_MinSquare = pO_MainShape.AreaIU
                        pO_MainShape.Delete
                        Set pO_MainShape = Shape
                        Application.ActiveWindow.Select pO_MainShape, visSelect
                    Else
                        VD_MinSquare = Shape.AreaIU
                        Shape.Delete
                    End If
                End If
            End If
    
    End Select

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "vO_App_ShapeAdded"
End Sub


'-----------------------------------Основные процедуры------------------------------------------------------------------------

Public Sub S_SetFullShape(ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Основная процедура запуска построения фигуры спргнозированного очага
Dim vO_Shape As Visio.Shape, vOTempShape As Visio.Shape
Dim v_X As Double, v_Y As Double
Dim i As Integer
Dim v_DoorsCount As Integer

On Error GoTo Tail

    Set col_FireShapes = New Collection
    
    '---Определяем коллекции фигур находящихся в опасной близости
        Set col_WallsShapes = New Collection
        Set col_DoorsShapes = New Collection
        Set col_PortalShapes = New Collection
        For Each vO_Shape In Application.ActivePage.Shapes
            '---Для коллекции дверей
            s_SearchDoors vO_Shape, a_X, a_Y, fD_MetersToInch(a_R)
        Next vO_Shape
        If col_DoorsShapes.Count > 0 Then
            ReDim arrDoorsIDArray(col_DoorsShapes.Count - 1)
        End If
        
    '---Очищаем значения остаточных радиусов для всех дверей в коллекции дверей нах в оп-й близости
        s_DoorsDistanceClear
    
    '---Формируем маску
        s_makeMask a_X, a_Y, fD_MetersToInch(a_R)
               
    '---Активируем дополнительные коолекции фигур
        Set col_DoorsFireShapes = New Collection
    
    '---Задаем стартовую позицию
    VI_PointsCount = 0
'    ReDim Col_Points(VI_PointsCount, 3)
    Col_Points(VI_PointsCount, 0) = a_X
    Col_Points(VI_PointsCount, 1) = a_Y
    Col_Points(VI_PointsCount, 2) = fD_MetersToInch(a_R)
    
    '---Строим прогноз развития пожара от затравочной точки
    For i = 0 To conI_TopPoint
        If Col_Points(i, 2) = 0 Then Exit For
        s_CircleSet Col_Points(i, 0), Col_Points(i, 1), Col_Points(i, 2)
    Next i
    
    '---Объединяем полученные фигуры
'    Debug.Print "Всего фигур: " & ActiveWindow.Selection.Count
    Debug.Print "Вторичных очагов: " & VI_PointsCount
    
'    '---Формируем окончательную фигуру
'    pO_MainFireShape.SendToBack
       
    '---Очищаем коллекцию площадей горения и массив точек для дальнейшего использования
'    Set col_FireShapes = New Collection
    Erase Col_Points()

    For v_DoorsCount = 0 To col_DoorsShapes.Count
        '---Строим прогноз развития пожара от дверей и проемов
'        For Each vO_Shape In col_DoorsFireShapes
        Set vO_Shape = Application.ActivePage.Shapes.ItemFromID(arrDoorsIDArray(v_DoorsCount))
        '---Проверяем, необходимо ли строить очаг от текущей двери
            If vO_Shape.Cells("Scratch.D1") <> 0 Then
            
            '---Задаем стартовую позицию - точка посередине внизу фигуры двери
                VI_PointsCount = 0
'                vO_Shape.XYToPage vO_Shape.Cells("Width").Result(visInches) / 2, -0.01, v_X, v_Y
                vO_Shape.XYToPage vO_Shape.Cells("Width").Result(visInches) / 2, 0, v_X, v_Y
                Col_Points(VI_PointsCount, 0) = v_X
                Col_Points(VI_PointsCount, 1) = v_Y
                Col_Points(VI_PointsCount, 2) = vO_Shape.Cells("Scratch.D1")
            '---Строим прогноз развития пожара от затравочной точки
                For i = 0 To conI_TopPoint
                    If Col_Points(i, 2) = 0 Then Exit For
                    s_CircleSet Col_Points(i, 0), Col_Points(i, 1), Col_Points(i, 2)
                Next i
                
            '---Задаем стартовую позицию - точка посередине вверху фигуры двери
                VI_PointsCount = 0
'                vO_Shape.XYToPage vO_Shape.Cells("Width").Result(visInches) / 2, vO_Shape.Cells("Height").Result(visInches) * 1.01, v_X, v_Y
                vO_Shape.XYToPage vO_Shape.Cells("Width").Result(visInches) / 2, vO_Shape.Cells("Height").Result(visInches) * 1, v_X, v_Y
                Col_Points(VI_PointsCount, 0) = v_X
                Col_Points(VI_PointsCount, 1) = v_Y
                Col_Points(VI_PointsCount, 2) = vO_Shape.Cells("Scratch.D1")
            '---Строим прогноз развития пожара от затравочной точки
                For i = 0 To conI_TopPoint
                    If Col_Points(i, 2) = 0 Then Exit For
                    s_CircleSet Col_Points(i, 0), Col_Points(i, 1), Col_Points(i, 2)
                Next i
            '---Указываем чрезмерно высокое значение остаточного радиуса для активной двери, для предотвращения повторного построения от нее
'                vO_Shape.Cells("Scratch.D1").Formula = 10000

            '---Вбрасываем фигуру площади дваери
                vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), v_X, v_Y
                Set vOTempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), 0, 0)
                vOTempShape.Cells("PinX").FormulaForce = v_X
                vOTempShape.Cells("PinY").FormulaForce = v_Y
                vOTempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
                Set pO_MainFireShape = pFO_Union(pO_MainFireShape, vOTempShape)

            End If
'        Next vO_Shape
    Next v_DoorsCount
    
    '---Перемещаем окончательную фигуру назад
    pO_MainFireShape.SendToBack
    
    '---Выделяем фигуру площади горения
        Application.ActiveWindow.Select pO_MainFireShape, visSelect
    
Set col_FireShapes = Nothing
Set col_DoorsShapes = Nothing
Set col_PortalShapes = Nothing

Exit Sub
Tail:
'    Debug.Print Err.Description
    Set col_WallsShapes = Nothing
    Set col_FireShapes = Nothing
    Set col_DoorsShapes = Nothing
    Set col_PortalShapes = Nothing
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "S_SetFullShape"
End Sub

Private Sub s_CircleSet(ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
Dim vO_Shape As Visio.Shape, vO_ActiveShape As Visio.Shape, vO_TempShape As Visio.Shape, vO_Canal As Visio.Shape
Dim Col_TempPoints(50, 2) As Double 'Временный массив координат критических точек
Dim x1 As Double, y1 As Double
Dim i As Integer, j As Integer
Dim vD_Distance As Double
Dim tx As Double, ty As Double, tr As Double

On Error GoTo Tail
    
    '---Запоминаем базовые координаты очага
    x = a_X
    y = a_Y
    
    '---Рисуем новый круг с указанными координатами
    Set vO_ActiveShape = Application.ActiveWindow.Page.DrawOval(a_X - a_R, a_Y - a_R, a_X + a_R, a_Y + a_R)
                  
    '---Вбрасываем тень маски
    Set vO_TempShape = Application.ActivePage.Drop(pO_MaskShape, 0, 0)
    vO_TempShape.Cells("PinX") = pO_MaskShape.Cells("PinX")
    vO_TempShape.Cells("PinY") = pO_MaskShape.Cells("PinY")
    
    '---Осуществляем процедуру Вычитания из Затравочной фигуры временно-вброшенной фигуры
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select vO_ActiveShape, visSelect
    Application.ActiveWindow.Select vO_TempShape, visSelect
    Application.ActiveWindow.Selection.Subtract
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    Set vO_ActiveShape = Application.ActiveWindow.Selection(1)
    
    '---Обрезаем от полученной фигуры лишнее ПРОБНОЕ
        Set vO_ActiveShape = pFO_Substract(pO_MainFireShape, vO_ActiveShape)
        If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
    '---Фрагментируем полученную фигуру для отброса отрезанных фигур
    VI_ShapesCheckFlag = 1
    '--- Определяем площадь активной фигуры для предотвращения некорректной обрезки
        VD_ShapeArea = vO_ActiveShape.AreaIU
        Application.ActiveWindow.Selection.Fragment
    VI_ShapesCheckFlag = 0
    vB_ShapeAlreadyDetectFlag = False
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    Set vO_ActiveShape = Application.ActiveWindow.Selection(1)

    '---Обрезаем от полученной фигуры лишнее ВЕРНОЕ
'        Set vO_ActiveShape = pFO_Substract(pO_MainFireShape, vO_ActiveShape)
'        If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    
    '---Отбрасываем слепые зоны
        s_CutBlind vO_ActiveShape, a_R
        If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
        Set vO_ActiveShape = Application.ActiveWindow.Selection(1)
    
    '!!!Проверяем наличие фигур дверей касающихся данной фигуры очага
        s_AddDoorsFires vO_ActiveShape, x, y, a_R

    '--- перебираем все строки в Geometry1
        i = 2
        Do While vO_ActiveShape.CellsSRCExists(visSectionFirstComponent, i, 0, 0) = True
'            X1 = vO_ActiveShape.Cells("PinX").Result(visInches) - (vO_ActiveShape.Cells("LocPinX").Result(visInches) - vO_ActiveShape.CellsSRC(visSectionFirstComponent, i, 0).Result(visInches))
'            Y1 = vO_ActiveShape.Cells("PinY").Result(visInches) - (vO_ActiveShape.Cells("LocPinY").Result(visInches) - vO_ActiveShape.CellsSRC(visSectionFirstComponent, i, 1).Result(visInches))
            ps_GetFirstPnt vO_ActiveShape, x1, y1, i, 0
            vD_Distance = Sqr((a_X - x1) ^ 2 + (a_Y - y1) ^ 2)
            '---в случае, если дистанция до точки меньше чем 0,99 радиуса и больше 0,05, то запускаем рекурсиный вброс вторичного очага
            If vD_Distance < a_R * 0.99 And vD_Distance > a_R * 0.02 Then
                '---Запоминаем координаты и радиус вторичного очага во временном массиве
                Col_TempPoints(j, 0) = x1
                Col_TempPoints(j, 1) = y1
                Col_TempPoints(j, 2) = a_R - vD_Distance
                j = j + 1
            End If
            i = i + 1
        Loop
        
        '---Работаем с временным массивом
            '---Сортируем массив по уменьшению в зависимости от радиуса
            For i = 0 To 50
                If Col_TempPoints(i, 2) = 0 Then Exit For
                For j = i To 50
                    If Col_TempPoints(j, 2) = 0 Then Exit For
                    If Col_TempPoints(i, 2) < Col_TempPoints(j, 2) Then
                        tx = Col_TempPoints(i, 0)
                        ty = Col_TempPoints(i, 1)
                        tr = Col_TempPoints(i, 2)
                        Col_TempPoints(i, 0) = Col_TempPoints(j, 0)
                        Col_TempPoints(i, 1) = Col_TempPoints(j, 1)
                        Col_TempPoints(i, 2) = Col_TempPoints(j, 2)
                        Col_TempPoints(j, 0) = tx
                        Col_TempPoints(j, 1) = ty
                        Col_TempPoints(j, 2) = tr
                    End If
                Next j
            Next i
            
            '---Сохраняем точки в главном массиве
            For i = 0 To 50
                If Col_TempPoints(i, 2) = 0 Then Exit For
                If fD_GetRadiuss(Col_TempPoints(i, 0), Col_TempPoints(i, 1)) = 0 Then 'если запись о данной точке отсутствует - _
                                                                                      'необходимо ее добавить!
                    VI_PointsCount = VI_PointsCount + 1
                    Col_Points(VI_PointsCount, 0) = Col_TempPoints(i, 0)
                    Col_Points(VI_PointsCount, 1) = Col_TempPoints(i, 1)
                    Col_Points(VI_PointsCount, 2) = Col_TempPoints(i, 2)
                ElseIf fD_GetRadiuss(x1, y1) < a_R - vD_Distance Then 'если запись о точке имеется, но _
                        'указанный радиус меньше текущего - необходимо его увеличить!
                    sD_SetRadiuss Col_TempPoints(i, 0), Col_TempPoints(i, 1), Col_TempPoints(i, 2)
    '                    Else 'указанная запись присутствует и текущий радиус меньше указанного - ничего не происходит!
                End If
            Next i
        
        
        
    '---Объединяем полученные фигуры
        Set pO_MainFireShape = pFO_Union(pO_MainFireShape, vO_ActiveShape)
        
Set vO_ActiveShape = Nothing
Set vO_Shape = Nothing

Exit Sub
Tail:
'   Debug.Print Err.Description
    Set vO_ActiveShape = Nothing
    Set vO_Shape = Nothing
'    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CircleSet"
End Sub






Private Sub s_CutBlind(ByRef aO_ActiveShape As Visio.Shape, ByVal a_R As Double)
'Процедура обрезки лишних секторов полученного круга с основаниями в точках изменения геометрии (вторичных очагах)
Dim vO_Shape As Visio.Shape
Dim X0 As Double, Y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
Dim vs_Angle As Single
Dim col_AddedShapes As Collection 'коллекция обрабатываемых фигур
Dim i As Integer, j As Integer
Dim vD_Distance As Double

Set col_AddedShapes = New Collection

On Error GoTo Tail

'--- добавляем активную фигуру в коллекцию обрабатываемых фигур
    col_AddedShapes.Add aO_ActiveShape
    
'--- Определяем площадь активной фигуры для предотвращения некорректной обрезки
    VD_ShapeArea = aO_ActiveShape.AreaIU
    
'--- определяем смещение координат
    X0 = x
    Y0 = y
    
    j = 0
    Do While aO_ActiveShape.CellsSRCExists(visSectionFirstComponent + j, i, 0, 0) = True
    '--- перебираем все строки в Geometry1 + j
        i = 2
        Do While aO_ActiveShape.CellsSRCExists(visSectionFirstComponent + j, i, 0, 0) = True
    
            '--- определить первую координату линии
'                X1 = aO_ActiveShape.Cells("PinX").Result(visInches) - (aO_ActiveShape.Cells("LocPinX").Result(visInches) - aO_ActiveShape.CellsSRC(visSectionFirstComponent + j, i, 0).Result(visInches))
'                Y1 = aO_ActiveShape.Cells("PinY").Result(visInches) - (aO_ActiveShape.Cells("LocPinY").Result(visInches) - aO_ActiveShape.CellsSRC(visSectionFirstComponent + j, i, 1).Result(visInches))
                ps_GetFirstPnt aO_ActiveShape, x1, y1, i, j
                
                vD_Distance = Sqr((X0 - x1) ^ 2 + (Y0 - y1) ^ 2)
                '---в случае, если дистанция до точки меньше чем 0,99 радиуса и больше 0,05, то строим линию отсечения
                If vD_Distance < a_R * 0.99 And vD_Distance > a_R * 0.05 Then
                '--- определить угол линии
                    vs_Angle = CalculateAngle(x1 - X0, y1 - Y0)
                '--- определить вторую координату линии
                    x2 = a_R * 1.05 * Cos(vs_Angle * (3.14159265358979 / 180)) + X0
                    y2 = a_R * 1.05 * Sin(vs_Angle * (3.14159265358979 / 180)) + Y0
                '--- построить по имеющимся координатам линию и добавляем ее в коллекцию обрабатываемых фигур
                    col_AddedShapes.Add Application.ActiveWindow.Page.DrawLine(x1, y1, x2, y2)
                End If
            i = i + 1
        Loop
        i = 0
        j = j + 1
    Loop
'--- обрезаем фигуру по построенным линиям
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_AddedShapes
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    '---Фрагментируем полученное выделение фигур для отброса отрезанных фигур
        VI_ShapesCheckFlag = 2
            Application.ActiveWindow.Selection.Fragment
        VI_ShapesCheckFlag = 0
        vB_ShapeAlreadyDetectFlag = False

Set col_AddedShapes = Nothing

Exit Sub

'Обработчки ошибок
Tail:
'    Debug.Print Err.Description
    Set col_AddedShapes = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CutBlind"
End Sub




Private Sub s_CutExisted(ByRef aO_ActiveShapeID As Long)
'Процедура вычитания из целевой фигуры фигур уже добавленных в коллекцию col_FireShapes
Dim vO_ColShape As Visio.Shape, vO_TempShape As Visio.Shape, vO_ActiveShape As Visio.Shape

    Set vO_ActiveShape = Application.ActivePage.Shapes.ItemFromID(aO_ActiveShapeID)
    For Each vO_ColShape In col_FireShapes
        If Not vO_ColShape = vO_ActiveShape Then
            Set vO_TempShape = Application.ActivePage.Drop(vO_ColShape, vO_ColShape.Cells("PinX").Result(visInches), vO_ColShape.Cells("PinY").Result(visInches))
    
            Application.ActiveWindow.DeselectAll
            Application.ActiveWindow.Select vO_ActiveShape, visSelect
            Application.ActiveWindow.Select vO_TempShape, visSelect
            Application.ActiveWindow.Selection.Subtract
    
            If Application.ActiveWindow.Selection.Count > 0 Then
                Set vO_ActiveShape = Application.ActiveWindow.Selection(1)
'                Exit Sub
            Else
                Set vO_ActiveShape = Nothing
                Set vO_ColShape = Nothing
                Set vO_TempShape = Nothing
                Exit Sub
            End If
        End If
    Next vO_ColShape

    Set vO_ActiveShape = Nothing
    Set vO_ColShape = Nothing
    Set vO_TempShape = Nothing
Exit Sub
EX:
    Set vO_ActiveShape = Nothing
    Set vO_ColShape = Nothing
    Set vO_TempShape = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CutExisted"
End Sub


Private Sub sD_SetRadiuss(ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура устанавливает значение радиуса в указанной точке
Dim i As Integer
Const vS_Tolerance = 0.0003

For i = 0 To conI_TopPoint

'---Проверяем, не соответствует ли координита в массиве указанной пользователем
    If Col_Points(i, 0) < a_X + vS_Tolerance And Col_Points(i, 0) > a_X - vS_Tolerance And _
        Col_Points(i, 1) < a_Y + vS_Tolerance And Col_Points(i, 1) > a_Y - vS_Tolerance Then
            Col_Points(i, 2) = a_R
            Exit Sub
    End If
Next i


End Sub


'-----------------------------------------Процедуры работы с фигурами дверей----------------------------------------------------
Private Sub s_AddDoorsFires(ByRef aL_FireShape As Visio.Shape, ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура проверяет все фигуры дверей находящихся в опасной близости от очага и в случае, если фигура двери касается
'указанного сегмента пожара, задаем для нее остаточный радиус
Dim vO_Door As Visio.Shape
Dim vS_ProtectedRadiussByDoor As Single 'Радиус защищенный дверью
Dim vS_LeftRadiuss As Single 'Остаточный радиус распространения огня от двери
Dim v_Distance As Double
Dim X0 As Double, Y0 As Double
    
    On Error GoTo EX
    For Each vO_Door In col_DoorsShapes
    '---Вычисляем защищенный радиус
        vS_ProtectedRadiussByDoor = PI_DoorsREI * fD_MetersToInch(PS_LineSpeedM)
        '---Проверяем, касается ли фигура двери фигуры данного сегмента очага...
        If fB_DoorTouchTest(vO_Door.Shapes(1), aL_FireShape) Then
            '...и если касается проверяем ее остаток огнестойкости, если он < REI то...
            vO_Door.XYFromPage a_X, a_Y, X0, Y0
            v_Distance = vO_Door.Shapes(1).DistanceFromPoint(X0, Y0, visSpatialIncludeDataGraphics) + _
                         Sqr(vO_Door.Cells("Height").Result(visInches) ^ 2 + (vO_Door.Cells("Width").Result(visInches) / 2) ^ 2)
'            v_Distance = Sqr((X0 - X1) ^ 2 + (Y0 - Y1) ^ 2)
            '---Если дистанция от двери + Защищенный дверью радиус, меньше радиуса пожара от указанной точки сегмента очага, то..
            If v_Distance + vS_ProtectedRadiussByDoor < a_R Then
            '---Устанавливаем в ячейку "Scratch.D1" значение остаточного радиуса развития пожара от фигуры двери
                'при условии что имеющееся в этой ячейке значение меньше нового
                vS_LeftRadiuss = a_R - (v_Distance + vS_ProtectedRadiussByDoor)
                If vO_Door.Cells("Scratch.D1") < vS_LeftRadiuss Then
                    vO_Door.Cells("Scratch.D1").Formula = vS_LeftRadiuss
                End If
            End If
            '---В случае, если фигура еще не была добавлена в коллекцию col_DoorsFireShapes - добавляем
'!!!!!!!!!!!Довести до ума!!!!!!!!!!!!!!1
            If fB_DoorsFireShapesExistsCheck(vO_Door) = False Then
                col_DoorsFireShapes.Add vO_Door
                s_addNewDoorID vO_Door.ID
                v_DoorsCount = v_DoorsCount + 1
            End If
        End If
    Next vO_Door

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_AddDoorsFires"
End Sub


Private Sub s_DoorsDistanceClear()
'Процедура очищает значения остаточного радиуса для всех дверей в коллекции дверей находящихся в опасной близости
Dim vO_shp As Visio.Shape

For Each vO_shp In col_DoorsShapes
    vO_shp.Cells("Scratch.D1").Formula = 0
Next vO_shp

Set vO_shp = Nothing
End Sub

Private Sub s_addNewDoorID(ByVal aI_DoorID As Integer)
'Процедура добавляет новое значение ID двери в массив
Dim i As Integer

    i = 0
    Do While arrDoorsIDArray(i) > 0
        i = i + 1
    Loop
    arrDoorsIDArray(i) = aI_DoorID

End Sub


'-----------------------------------------Процедуры определения коллекций-------------------------------------------------------
Private Sub s_SearchWalls(ByRef aO_Shape As Visio.Shape, ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_WallsShapes СТЕНЫ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape
Dim v_X As Double, v_Y As Double

    On Error GoTo EX
'---Проверяем, находится ли фигура в зоне досягаемости очага
    My_XYFromPage aO_Shape, a_X, a_Y, v_X, v_Y
    If aO_Shape.HitTest(v_X, v_Y, a_R * 1.05) = 0 Then Exit Sub

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            s_SearchWalls vO_shp, a_X, a_Y, a_R
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        Exit Sub
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        col_WallsShapes.Add aO_Shape
        Exit Sub
    End If

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_SearchWalls"
End Sub

Private Sub s_SearchDoors(ByRef aO_Shape As Visio.Shape, ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_DoorsShapes ДВЕРИ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape
Dim v_X As Double, v_Y As Double

    On Error GoTo EX
'---Проверяем, находится ли фигура в зоне досягаемости очага
    My_XYFromPage aO_Shape, a_X, a_Y, v_X, v_Y
    If aO_Shape.HitTest(v_X, v_Y, a_R * 1.05) = 0 Then Exit Sub

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            s_SearchDoors vO_shp, a_X, a_Y, a_R
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        Exit Sub
    End If
    
'---Проверяем, является ли фигура фигурой ДВЕРЬ
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Then
        col_DoorsShapes.Add aO_Shape
        Exit Sub
    End If

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_SearchDoors"
End Sub


Private Sub s_SearchPortals(ByRef aO_Shape As Visio.Shape, ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура рекурсивно перебирает ВСЕ фигуры находящиеся на листе и добавляет в коллекцию col_PortalShapes ПРОХОДЫ находящиеся
'в пределах первичной площади
Dim vO_shp As Visio.Shape
Dim v_X As Double, v_Y As Double

    On Error GoTo EX
'---Проверяем, находится ли фигура в зоне досягаемости очага
    My_XYFromPage aO_Shape, a_X, a_Y, v_X, v_Y
    If aO_Shape.HitTest(v_X, v_Y, a_R * 1.05) = 0 Then Exit Sub

'---Проверяем, является ли фигура составной (сгруппированной)
    If aO_Shape.Shapes.Count > 1 Then
        For Each vO_shp In aO_Shape.Shapes
            s_SearchPortals vO_shp, a_X, a_Y, a_R
        Next vO_shp
        Set vO_shp = Nothing
    End If

'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        Exit Sub
    End If
    
'---Проверяем, является ли фигура фигурой ПРОХОД
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25 Then
        col_PortalShapes.Add aO_Shape
        Exit Sub
    End If

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_SearchPortals"
End Sub

'-----------------------------------------------------Процедуры создания маски----------------------------------------------------

Private Sub s_makeMask(ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double)
'Процедура создает маску обрезки (ОСНОВНАЯ)
Dim vO_Shape As Visio.Shape

    On Error GoTo EX
    '---Формируем коллекции
    For Each vO_Shape In Application.ActivePage.Shapes
        '---Для коллекции стен
        s_SearchWalls vO_Shape, a_X, a_Y, a_R
        '---Для коллекции проходов
        s_SearchPortals vO_Shape, a_X, a_Y, a_R
        '---Для коллекции дверей Делается предварительно!!!
'        s_SearchDoors vO_Shape
'        ReDim arrDoorsIDArray(col_DoorsShapes.Count)
        
    Next vO_Shape

    '---Делаем маску стен
   s_makeMaskWalls

    '---Добавляем в маску проходы
    s_CorrectMaskByPortals
    
    '---Добавляем в маску дверные проемы (прми REI = 0 или открытых дверях)
        s_CorrectMaskByDoors
        ReDim arrDoorsIDArray(col_DoorsShapes.Count)

Set vO_Shape = Nothing
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_makeMask"
End Sub



Private Sub s_makeMaskWalls()
'Процедура создает маску обрезки стен
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---Проверяем имеются ли фигуры в коллекции col_WallsShapes
    If col_WallsShapes.Count = 0 Then Exit Sub

'---перебираем все фигуры стен
    For Each vO_Shape In col_WallsShapes
        '---Определяем координаты для вброса
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
'            My_XYToPage vO_Shape, vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), X1, Y1
        '---Вбрасываем новую фигуру по координатам
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
            vO_TempShape.Cells("PinX") = x1
            vO_TempShape.Cells("PinY") = y1
            vO_TempShape.Cells("LocPinX") = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY") = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("Angle").FormulaForce = AngleToPage(vO_Shape)
'            vO_TempShape.Cells("Angle").FormulaU = AngleToPage(vO_Shape) & "deg"
            
        '---Добавляем во временную коллекцию
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---Формируем на основе собранной коллекции маску
    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Union
    Application.ActiveWindow.Selection.SendToBack

'---Запоминаем маску
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

Exit Sub

Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_makeMaskWalls"
End Sub

Private Sub s_CorrectMaskByPortals()
'Процедура добавляет в фигуру маски проходы
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---Проверяем имеются ли фигуры в коллекции col_PortalShapes
    If col_PortalShapes.Count = 0 Then Exit Sub

'---перебираем все фигуры проходов
    For Each vO_Shape In col_PortalShapes
        '---Определяем координаты для вброса
            vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
'            My_XYToPage vO_Shape, vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), X1, Y1
        '---Вбрасываем новую фигуру по координатам
            Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
            vO_TempShape.Cells("PinX").FormulaForce = x1
            vO_TempShape.Cells("PinY").FormulaForce = y1
            vO_TempShape.Cells("LocPinX").FormulaForce = vO_Shape.Cells("LocPinX")
            vO_TempShape.Cells("LocPinY").FormulaForce = vO_Shape.Cells("LocPinY")
            vO_TempShape.Cells("Angle").Formula = AngleToPage(vO_Shape)
'            vO_TempShape.Cells("Angle").FormulaForceU = AngleToPage(vO_Shape) & "deg"
            
        '---Добавляем во временную коллекцию
            col_Shadows.Add vO_TempShape
    Next vO_Shape

'---Создаем временную тень маски
    Set vO_TempShape = Application.ActivePage.Drop(pO_MaskShape, x1, y1)
    vO_TempShape.Cells("PinX") = pO_MaskShape.Cells("PinX")
    vO_TempShape.Cells("PinY") = pO_MaskShape.Cells("PinY")

'---Корректируем при помощи собранной коллекции маску - удаляем из не тени проходов
'    Application.ActiveWindow.DeselectAll
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Subtract
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    Application.ActiveWindow.Selection.SendToBack

'---Запоминаем маску
    pO_MaskShape.Delete
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

Exit Sub

Tail:
    Debug.Print Err.description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CorrectMaskByPortals"


End Sub

Private Sub s_CorrectMaskByDoors()
'Процедура добавляет в фигуру маски дверные проемы, в случае, если огнестойкость открытых дверей = 0
'или PB_CheckOpens = True
Dim vO_Shape As Visio.Shape
Dim vO_TempShape As Visio.Shape
Dim col_Shadows As Collection
Dim x1 As Double, y1 As Double

Set col_Shadows = New Collection

On Error GoTo Tail

'---Проверяем имеются ли фигуры в коллекции col_DoorsShapes
    If col_DoorsShapes.Count = 0 Then Exit Sub

'---перебираем все фигуры проходов
    For Each vO_Shape In col_DoorsShapes
        '---Проверяем на соответствие требованиям
            If fB_CheckDoorsTransparency(vO_Shape) = True Then
                '---Определяем координаты для вброса
                    vO_Shape.XYToPage vO_Shape.Cells("LocPinX").Result(visInches), vO_Shape.Cells("LocPinY").Result(visInches), x1, y1
                '---Вбрасываем новую фигуру по координатам
                    Set vO_TempShape = Application.ActivePage.Drop(vO_Shape.Shapes(1), x1, y1)
                    vO_TempShape.Cells("PinX").FormulaForce = x1
                    vO_TempShape.Cells("PinY").FormulaForce = y1
                    vO_TempShape.Cells("LocPinX").FormulaForce = vO_Shape.Cells("LocPinX")
                    vO_TempShape.Cells("LocPinY").FormulaForce = vO_Shape.Cells("LocPinY")
                    vO_TempShape.Cells("Angle").Formula = AngleToPage(vO_Shape)
                    
                '---Добавляем во временную коллекцию
                    col_Shadows.Add vO_TempShape
            End If
    Next vO_Shape

'---Создаем временную тень маски
    Set vO_TempShape = Application.ActivePage.Drop(pO_MaskShape, x1, y1)
    vO_TempShape.Cells("PinX") = pO_MaskShape.Cells("PinX")
    vO_TempShape.Cells("PinY") = pO_MaskShape.Cells("PinY")

'---Корректируем при помощи собранной коллекции маску - удаляем из неё тени дверных проемов
    For Each vO_Shape In col_Shadows
        Application.ActiveWindow.Select vO_Shape, visSelect
    Next vO_Shape
    Application.ActiveWindow.Selection.Subtract
    If Application.ActiveWindow.Selection.Count = 0 Then Exit Sub
    Application.ActiveWindow.Selection.SendToBack

'---Запоминаем маску
    pO_MaskShape.Delete
    Set pO_MaskShape = Application.ActiveWindow.Selection(1)

'---Удаляем из коллекции дверей двери подпавшие под исходный критерий (огнестойкость открытых дверей = 0
    'или PB_CheckOpens = True
    '---перебираем все фигуры проходов
        For Each vO_Shape In col_DoorsShapes
            For Each vO_TempShape In col_Shadows
                If vO_Shape = vO_TempShape Then
                    col_DoorsShapes.Remove vO_Shape.Name
                    col_Shadows.Remove vO_TempShape.Name
                End If
            Next vO_TempShape
        Next vO_Shape

Exit Sub

Tail:
'    Debug.Print Err.Description
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CorrectMaskByDoors"


End Sub



Private Sub My_XYToPage(ByRef a_shp As Visio.Shape, ByVal a_InX As Double, ByVal a_InY As Double, _
                        ByRef a_OutX As Double, ByRef a_OutY As Double)
'Процедура корректного перевода координат к системе координат страницы. На вход фигура в чьей системе изначальные координаты,
'сами координаты в системе материнской фигуры.
'А так же ссылка на переменные необходимые для изменения

If a_shp.Parent.Name = Application.ActivePage.Name Then
    a_OutX = a_InX
    a_OutY = a_InY
Else
    a_shp.XYToPage a_InX, a_InY, a_OutX, a_OutY
End If


End Sub

Private Sub My_XYFromPage(ByRef a_shp As Visio.Shape, ByVal a_InX As Double, ByVal a_InY As Double, _
                        ByRef a_OutX As Double, ByRef a_OutY As Double)
'Процедура корректного перевода координат к системе координат для фигуры. На вход фигура в чью систему необходимо перевести координаты,
'сами координаты в системе страницы.
'А так же ссылка на переменные необходимые для изменения

If a_shp.Parent.Name = Application.ActivePage.Name Then
    a_OutX = a_InX
    a_OutY = a_InY
Else
    a_shp.Parent.XYFromPage a_InX, a_InY, a_OutX, a_OutY
End If


End Sub


'-----------------------------------------Procs класса------------------------------------------------------------------------
Private Sub ps_GetFirstPnt(ByRef a_shp As Visio.Shape, ByRef a_X As Double, ByRef a_Y As Double, ByVal i As Integer, ByVal j As Integer)
'Proc for getting page coordinates of first point

    Select Case Int(Application.version)
        Case Is = 15
            a_shp.XYToPage a_shp.CellsSRC(visSectionFirstComponent + j, i, 0).Result(visNumber) * a_shp.Cells("Width").Result(visInches), _
                a_shp.CellsSRC(visSectionFirstComponent + j, i, 1).Result(visNumber) * a_shp.Cells("Height").Result(visInches), a_X, a_Y
        Case Is = 16
            a_shp.XYToPage a_shp.CellsSRC(visSectionFirstComponent + j, i, 0).Result(visNumber) * a_shp.Cells("Width").Result(visInches), _
                a_shp.CellsSRC(visSectionFirstComponent + j, i, 1).Result(visNumber) * a_shp.Cells("Height").Result(visInches), a_X, a_Y

        Case Else
            a_shp.XYToPage a_shp.CellsSRC(visSectionFirstComponent + j, i, 0).Result(visInches), _
                a_shp.CellsSRC(visSectionFirstComponent + j, i, 1).Result(visInches), a_X, a_Y
    End Select

End Sub



'-----------------------------------------Функции класса------------------------------------------------------------------------

Private Function CalculateAngle(ByVal x As Double, ByVal y As Double) As Single 'в градусах
'Функция возвращает значение угла (в градусах) прямой проведенной из начала координат к точке с указанными координатами
    If x > 0 And y = 0 Then
        CalculateAngle = 0
    ElseIf x = 0 And y > 0 Then
        CalculateAngle = 90
    ElseIf x < 0 And y = 0 Then
        CalculateAngle = 180
    ElseIf x = 0 And y < 0 Then
        CalculateAngle = 270
    ElseIf x > 0 And y > 0 Then
        CalculateAngle = Atn(y / x) * (180 / 3.14159265358979)
    ElseIf x < 0 And y > 0 Then
        CalculateAngle = Atn(y / x) * (180 / 3.14159265358979) + 180
    ElseIf x > 0 And y < 0 Then
        CalculateAngle = Atn(y / x) * (180 / 3.14159265358979)
    ElseIf x < 0 And y < 0 Then
        CalculateAngle = Atn(y / x) * (180 / 3.14159265358979) - 180
    End If
End Function


Private Function s_CheckConsists(ByVal aO_ActiveShapeID As Long) As Boolean
'Функция проверки нахождения указанной фигуры внутри уже вброшенных фигур очага пожара
Dim vO_ColShape As Visio.Shape, vO_ActiveShape As Visio.Shape
Dim x1, y1 As Double
Dim i As Integer
Dim vB_PointFlag As Boolean

    On Error GoTo EX
'---Если коллекция col_FireShapes пуста, то функция возвращает False, т.к. это автоматически означает, что мы имеем дело с первой фигурой
    If col_FireShapes.Count = 0 Then
        s_CheckConsists = False
        Exit Function
    End If

    Set vO_ActiveShape = Application.ActivePage.Shapes.ItemFromID(aO_ActiveShapeID)
    vB_PointFlag = False

    '--- перебираем все строки в Geometry1
        i = 2
        Do While vO_ActiveShape.CellsSRCExists(visSectionFirstComponent, i, 0, 0) = True
        '--- перебираем все фигуры в коллекции площаде пожара
            x1 = vO_ActiveShape.Cells("PinX").Result(visInches) - (vO_ActiveShape.Cells("LocPinX").Result(visInches) - vO_ActiveShape.CellsSRC(visSectionFirstComponent, i, 0).Result(visInches))
            y1 = vO_ActiveShape.Cells("PinY").Result(visInches) - (vO_ActiveShape.Cells("LocPinY").Result(visInches) - vO_ActiveShape.CellsSRC(visSectionFirstComponent, i, 1).Result(visInches))
                
            For Each vO_ColShape In col_FireShapes
                If vO_ColShape.HitTest(x1, y1, 0.01) > 0 Then 'одна из точек фигуры выступает за контур вдрошенных фигур площади
                    vB_PointFlag = True
                End If
            Next vO_ColShape
            
            If vB_PointFlag = False Then
                s_CheckConsists = False
                Set vO_ActiveShape = Nothing
                Set vO_ColShape = Nothing
                Exit Function
            End If
            vB_PointFlag = False
            i = i + 1
        Loop
    
s_CheckConsists = True

Set vO_ActiveShape = Nothing
Set vO_ColShape = Nothing
Exit Function
EX:
    Set vO_ActiveShape = Nothing
    Set vO_ColShape = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "s_CheckConsists"
End Function

Private Function f_LayInTest(ByVal aO_FirstShapeID As Long, ByVal aO_SecondShapeID As Long) As Boolean
'Функция определят находится ли первая фигура внутри второй
Dim i As Integer
Dim x1 As Double, y1 As Double
Dim vO_FirstShape As Visio.Shape, vO_SecondShape As Visio.Shape

    Set vO_FirstShape = Application.ActivePage.Shapes.ItemFromID(aO_FirstShapeID)
    Set vO_SecondShape = Application.ActivePage.Shapes.ItemFromID(aO_SecondShapeID)

'--- перебираем все строки в Geometry1 Первой фигуры
    i = 1
    Do While vO_FirstShape.CellsSRCExists(visSectionFirstComponent, i, 0, 0) = True
        x1 = vO_FirstShape.Cells("PinX").Result(visInches) - (vO_FirstShape.Cells("LocPinX").Result(visInches) - vO_FirstShape.CellsSRC(visSectionFirstComponent, i, 0).Result(visInches))
        y1 = vO_FirstShape.Cells("PinY").Result(visInches) - (vO_FirstShape.Cells("LocPinY").Result(visInches) - vO_FirstShape.CellsSRC(visSectionFirstComponent, i, 1).Result(visInches))
        '---в случае, если дистанция до точки меньшечем 0,99 радиуса и больше 0,05, то запускаем рекурсиный вброс вторичного очага
        If vO_SecondShape.HitTest(x1, y1, 0.01) = 0 Then
            f_LayInTest = False
            Exit Function
        End If
        i = i + 1
    Loop

f_LayInTest = True

End Function


Private Function fB_DoorTouchTest(ByRef aO_FirstShape As Visio.Shape, ByVal aO_SecondShape As Visio.Shape) As Boolean
'Функция возвращает True если фигура двери касается второй указанной фигуры и False если нет
Dim i As Integer
Dim v_X As Double, v_Y As Double

'--- перебираем все углы фигуры двери
    '---Левый нижний угол
        My_XYToPage aO_FirstShape, 0, 0, v_X, v_Y
        If aO_SecondShape.HitTest(v_X, v_Y, 0.01) > 0 Then
            fB_DoorTouchTest = True
            Exit Function
        End If
    '---Левый верхний угол
        My_XYToPage aO_FirstShape, 0, aO_FirstShape.Cells("Height").Result(visInches), v_X, v_Y
        If aO_SecondShape.HitTest(v_X, v_Y, 0.01) > 0 Then
            fB_DoorTouchTest = True
            Exit Function
        End If
    '---Правый верхний угол
        My_XYToPage aO_FirstShape, aO_FirstShape.Cells("Width").Result(visInches), _
                                                                aO_FirstShape.Cells("Height").Result(visInches), v_X, v_Y
        If aO_SecondShape.HitTest(v_X, v_Y, 0.01) > 0 Then
            fB_DoorTouchTest = True
            Exit Function
        End If
    '---Правый нижний угол
        My_XYToPage aO_FirstShape, aO_FirstShape.Cells("Width").Result(visInches), 0, v_X, v_Y
        If aO_SecondShape.HitTest(v_X, v_Y, 0.01) > 0 Then
            fB_DoorTouchTest = True
            Exit Function
        End If

fB_DoorTouchTest = False

'Set vO_FirstShape = Nothing
'Set vO_SecondShape = Nothing
End Function


Private Function fD_CheckRadiuss(ByVal a_X As Double, ByVal a_Y As Double, ByVal a_R As Double) As Double
'Функция проверяет наличие точки с указанными координатами в массиве координат известных точек_
'и если такая точка имеется, проверяет не превышает ли указанное значение радиуса указанное ранее_
'возвращает большее значение
Dim i As Integer
Const vS_Tolerance = 0.01

    For i = 0 To conI_TopPoint
    
    '---Проверяем, не соответствует ли координита в массиве указанной пользователем
        If Col_Points(i, 0) < a_X + vS_Tolerance And Col_Points(i, 0) > a_X - vS_Tolerance And _
            Col_Points(i, 1) < a_Y + vS_Tolerance And Col_Points(i, 1) > a_Y - vS_Tolerance Then
            If Col_Points(i, 2) < a_R Then
                Col_Points(i, 2) = a_R
                fD_CheckRadiuss = a_R
                Exit Function
            Else
                fD_CheckRadiuss = Col_Points(i, 2)
                Exit Function
            End If
        
        End If
    Next i

fD_CheckRadiuss = a_R
End Function


Private Function fD_GetRadiuss(ByVal a_X As Double, ByVal a_Y As Double) As Double
'Функция возвращает значение радиуса в указанной точке
Dim i As Integer
Const vS_Tolerance = 0.0003

    For i = 0 To conI_TopPoint
    '---Проверяем, не соответствует ли координита в массиве указанной пользователем
        If Col_Points(i, 0) < a_X + vS_Tolerance And Col_Points(i, 0) > a_X - vS_Tolerance And _
            Col_Points(i, 1) < a_Y + vS_Tolerance And Col_Points(i, 1) > a_Y - vS_Tolerance Then
                fD_GetRadiuss = Col_Points(i, 2)
                Exit Function
        End If
    Next i

fD_GetRadiuss = 0
End Function


Private Function fD_MetersToInch(ByVal aD_Meters As Double) As Double
'Процедура получает длину в метрах, а возвращает в дюймах
    fD_MetersToInch = aD_Meters * 39.37
End Function


Private Function AngleToPage(ByRef Shape As Visio.Shape) As Double
'Dim shp As Visio.Shape

'Set shp = Application.ActivePage.Shapes.ItemFromID(ShapeXYID)

If Shape.Parent.Name = Application.ActivePage.Name Then
    AngleToPage = Shape.Cells("Angle")
Else
'    AngleToPage = Shape.Cells("Angle").Result(visDegrees) + AngleToPage(Shape.Parent)
    AngleToPage = Shape.Cells("Angle") + AngleToPage(Shape.Parent)
End If

Set Shape = Nothing
End Function



Private Function fB_DoorsFireShapesExistsCheck(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция ищет указанную фигуру в коллекции, и возвращает True, если она там уже есть
'либо False, если нет
Dim vOshp_Shape As Visio.Shape

'Set vOshp_Shape = Application.ActivePage.Shapes.ItemFromID(aO_ShapeID)

    For Each vOshp_Shape In col_DoorsFireShapes
        If vOshp_Shape = aO_Shape Then
            fB_DoorsFireShapesExistsCheck = True
            Exit Function
        End If
    Next vOshp_Shape

fB_DoorsFireShapesExistsCheck = False

Set vOshp_Shape = Nothing
'Set vOshp_Shape = Nothing
End Function



Private Function pFO_Union(ByRef FstShape As Visio.Shape, ByRef SndShape As Visio.Shape) As Visio.Shape
'Функция возвращает ссылку на фигуру собранную из двух переданных

    On Error GoTo Tail

'    If FstShape = Empty Then
'        Set pFO_Union = SndShape
'        Exit Function
'    End If

    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select FstShape, visSelect
    Application.ActiveWindow.Select SndShape, visSelect
    Application.ActiveWindow.Selection.Union
    Set pFO_Union = Application.ActiveWindow.Selection(1)

Exit Function
Tail:
    Set pFO_Union = SndShape
End Function

Private Function pFO_Substract(ByRef FstShape As Visio.Shape, ByRef SndShape As Visio.Shape) As Visio.Shape
'Функция возвращает ссылку на фигуру полученную обрезанием первой от второй из двух переданных
Dim vOS_Shadow As Visio.Shape

    On Error GoTo Tail
    
    Set vOS_Shadow = Application.ActivePage.Drop(FstShape, 0, 0)
    vOS_Shadow.Cells("PinX") = FstShape.Cells("PinX")
    vOS_Shadow.Cells("PinY") = FstShape.Cells("PinY")
    
    Application.ActiveWindow.DeselectAll
    Application.ActiveWindow.Select SndShape, visSelect
    Application.ActiveWindow.Select vOS_Shadow, visSelect
    Application.ActiveWindow.Selection.Subtract
    Set pFO_Substract = Application.ActiveWindow.Selection(1)

Exit Function
Tail:
    Set pFO_Substract = SndShape
End Function

Private Function fB_CheckDoorsTransparency(ByRef aO_DoorShape As Visio.Shape) As Boolean
'Функция возвращает True если свойство PB_CheckOpens = True и при этом дверь открыта
'Так же функция возвращает True при REI = 0
'!!!На вход принимаются только фигуры дверей!!!
Dim vI_DoorType As Integer 'Тип двери

On Error GoTo Tail

    '---проверяем имеем ли мы дело с фигурой двери, если нет - выходим из функции
    If Not aO_DoorShape.Cells("User.ShapeType").Result(visNumber) = 10 Then
        fB_CheckDoorsTransparency = False
        Exit Function
    End If

    vI_DoorType = aO_DoorShape.Cells("User.SubShapeType").Result(visNumber)
    Select Case vI_DoorType
        Case Is = 74 'Одностворчатая, двустворчатая распашная
            If (PB_CheckOpens = True And aO_DoorShape.Shapes(9).Cells("Angle") <> 0) Or PI_DoorsREI = 0 Then
                fB_CheckDoorsTransparency = True
            End If
        Case Is = 28 'Двойная, С разными створками, Противолежащие створки
            Select Case aO_DoorShape.Cells("User.SubShapeType.Prompt").ResultStr(visUnitsString)
            Case Is = "Двойная"
                If (PB_CheckOpens = True And aO_DoorShape.Shapes(11).Cells("Angle") <> 0) Or PI_DoorsREI = 0 Then
                    fB_CheckDoorsTransparency = True
                End If
            Case Is = "Противолежащие створки"
                If (PB_CheckOpens = True And aO_DoorShape.Shapes(9).Cells("Angle") <> 0) Or PI_DoorsREI = 0 Then
                    fB_CheckDoorsTransparency = True
                End If
            Case Is = "С разными створками"
                If (PB_CheckOpens = True And aO_DoorShape.Shapes(11).Cells("Angle") <> 0) Or PI_DoorsREI = 0 Then
                    fB_CheckDoorsTransparency = True
                End If
            End Select
        Case Is = 64 'Вращающаяся
            If (PB_CheckOpens = True And _
                (aO_DoorShape.Shapes(9).Cells("Angle").ResultInt(visDegrees, 1) = 0 Or aO_DoorShape.Shapes(10).Cells("Angle").ResultInt(visDegrees, 1) = 0 _
                Or aO_DoorShape.Shapes(11).Cells("Angle").ResultInt(visDegrees, 1) = 0 Or aO_DoorShape.Shapes(12).Cells("Angle").ResultInt(visDegrees, 1) = 0)) _
                Or PI_DoorsREI = 0 Then
                fB_CheckDoorsTransparency = False
            Else
                fB_CheckDoorsTransparency = True
            End If
            
        '!!!Все прочие ПОКА рассчитываются обычным образом
        Case Is = 72 'С карманом
            fB_CheckDoorsTransparency = False
        Case Is = 26 'Двустворчатая с карманом
            fB_CheckDoorsTransparency = False
        Case Is = 73 'Одностворчатая скользящая
            fB_CheckDoorsTransparency = False
        Case Is = 71 'Одностворчатая складывающаяся
            fB_CheckDoorsTransparency = False
        Case Is = 25 'Двустворчатая складывающаяся
            fB_CheckDoorsTransparency = False
        Case Is = 65 'Рольставни
            fB_CheckDoorsTransparency = False
        Case Else    'Для всех прочих случаев
            fB_CheckDoorsTransparency = False
    End Select
    
Exit Function

Tail:
'    Debug.Print Err.Description
    fB_CheckDoorsTransparency = False
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "Sm_ShapeFormShow"
End Function


