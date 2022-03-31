VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Modeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------Класс моделирующий поведение пожара----------------------------------------------
Const diag As Double = 0.1                                  ' Коэффициент мощности по диагонали
Const orto As Double = 0.14142135623731                     ' Коэффициент мощности по вертикали и горизонтали
Const mmInInch As Double = 25.4

Private matrix As c_Matrix
Private extSquare As c_ExtinguishingSquare                  'Объект площади тушения

Public grain As Double              'размер зерна расчета (клетки)
Private step As Integer             'Счетчик количества шагов (общий, без моделирвоания и с моделирвоанием)
Private calcStep As Integer         'Счетчик количества раелизованных (ТОЛЬКО с моделированием) шагов
Public distance As Single           'общая дистанция пройденная при моделирвоании (условная для идеального случая)
Public time As Single               'общее время прошедшее при моделирвоании (условная для идеального случая)
Public intenseNeed As Single        'Требуемая интенсивность подачи воды л/(с*м.кв.)

Private nozzleDemon As c_FireNozzlesDemon
    


Public Sub SetMatrix(ByRef m As c_Matrix)
'Устанавливаем ссылку на матрицу
    Set matrix = m
End Sub

Public Sub refreshOpenSpaces(ByRef outerMatrix() As Variant)
'Обновляем матрицу открытых пространств
    matrix.SetOpenSpace outerMatrix
End Sub

Public Sub SetFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'Создаем новую точку горения
    matrix.SetFireCell x, y, value
    
    'Добавляем запись о горящей точке в коллекцию активных клеток
    AddToActiveCells x, y
    
    'Insert Firepoint as Visio Shape
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
End Sub

Public Sub SetStartFireCell(ByVal x As Long, ByVal y As Long, Optional value As Double)
'Создаем стартовую точку горения
    matrix.SetFireCell x, y, value
    'добавляем точки распространения горения
    matrix.SetFireCell x - 2, y - 2, 3.4
    matrix.SetFireCell x - 1, y - 2, 9.61
    matrix.SetFireCell x, y - 2, 13.59
    matrix.SetFireCell x + 1, y - 2, 9.61
    matrix.SetFireCell x + 2, y - 2, 3.4

    matrix.SetFireCell x - 2, y - 1, 9.61
    matrix.SetFireCell x - 1, y - 1, 43.58
    matrix.SetFireCell x, y - 1, 61.64
    matrix.SetFireCell x + 1, y - 1, 43.58
    matrix.SetFireCell x + 2, y - 1, 9.61

    matrix.SetFireCell x - 2, y, 13.59
    matrix.SetFireCell x - 1, y, 61.64
    matrix.SetFireCell x + 1, y, 61.64
    matrix.SetFireCell x + 2, y, 13.59

    matrix.SetFireCell x - 2, y + 1, 9.61
    matrix.SetFireCell x - 1, y + 1, 43.58
    matrix.SetFireCell x, y + 1, 61.64
    matrix.SetFireCell x + 1, y + 1, 43.58
    matrix.SetFireCell x + 2, y + 1, 9.61

    matrix.SetFireCell x - 2, y + 2, 3.4
    matrix.SetFireCell x - 1, y + 2, 9.61
    matrix.SetFireCell x, y + 2, 13.59
    matrix.SetFireCell x + 1, y + 2, 9.61
    matrix.SetFireCell x + 2, y + 2, 3.4
    
    'Добавляем запись о горящей точке в коллекцию активных клеток
    AddToActiveCells x, y
    
    'Insert Firepoint as Visio Shape
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
End Sub

Public Function AddFireFromShape(ByRef fireShape As Visio.Shape) As c_Modeller
'Добавление к площади пожара указанной фигуры
Dim fireShapeDemon As c_FireShapeDemon

    Set fireShapeDemon = New c_FireShapeDemon
    'Расчитываем область горения фигуры
    fireShapeDemon.setGrain(grain).SetFireShape(fireShape).RunDemon matrix

    Set fireShapeDemon = Nothing
    
Set AddFireFromShape = Me
End Function

Public Sub ActivateNozzles(ByRef frm As Object)
    Set nozzleDemon = New c_FireNozzlesDemon
    nozzleDemon.SetForm(frm).setGrain = grain
    Set nozzleDemon.SetMatrix = matrix
    nozzleDemon.FindNozzles
End Sub



Public Sub OneRound()
'один шаг моделирования
Dim cell As c_Cell
    
    'РАЗОБРАТЬСЯ С ЭТИМ - рисование дополнительного шага даже при поданном стволе!!!
'    'Запускаем анализ наличия пожарнных стволов
'    Dim nozzleDemon As c_FireNozzlesDemon
'    Set nozzleDemon = New c_FireNozzlesDemon
'    nozzleDemon.SetGrain = grain
'    Set nozzleDemon.setMatrix = matrix
'    nozzleDemon.RunDemon
'
'    'Возможно имеет смысл оставить в качестве опции (или еще как то) безусловный расчет площади тушения, в независимаости от того, поданы стволы или нет - это важно
'    'Рассчитываем площади тушения для каждого имеющегося на странице ствола, при условии, что для него укзана точка контакта с фронтом пожара
'    FireSquareCalculate nozzleDemon.nozzles
    
    
    
    
    
    'Очищаем коллекцию клеток только начинающих гореть (разгорающихся)
    Set matrix.startedFireCells = New Collection
    
    'Очищаем коллекцию ключевых точек
    Set matrix.keyCornerCells = New Collection

    'Атакуем всеми активными клетками (не внутренними)
    For Each cell In matrix.activeFiredCells
        Attack cell.x, cell.y
    Next cell
    
    '--------------------
    'После атаки увеличиваем значения горения всех клеток в соответствии с полученным ими энергиями
    For Each cell In matrix.chargedCells
        RizeCellFire cell.x, cell.y
    Next cell
    
    'Очищаем коллекцию клеток получивших заряд
    Set matrix.chargedCells = New Collection
    
    'Пока не удалять - этот блок здесь не спроста был...
'    'Проверяем клетки, на предмет того, не стали ли они внтуренними, и в этом случае удаляем их из коллекций
'    For Each cell In matrix.activeFiredCells
'        If matrix.IsInner(cell.x, cell.y) Then
'            RemoveFromActiveCells cell.x, cell.y
'            RemoveFromFrontCells cell.x, cell.y
'        End If
'        If matrix.IsNotPerimeter(cell.x, cell.y) Then
'            RemoveFromPerimeterCells cell.x, cell.y
'        End If
'    Next cell

    
    'Спрямляем горение у стен (проверяем все клетки, что получили энергию, но еще не горят полностью))
    If F_InsertFire.Flag_CornerDemonIsActive.value = True Then
        Dim shp As Visio.Shape
        Dim cornerDemon As c_CornerDemon
        Set cornerDemon = New c_CornerDemon
        cornerDemon.Init matrix, 0, 0
        For Each cell In matrix.fireFrontCells
            If matrix.IsTripleStateCell(cell.x, cell.y) Then
            
            '!!!Отрисовываем квадрат если клетка является ключевой
    
    '            Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
    '                    (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
    '            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Угловые точки")
    '            shp.Cells("LinePattern").FormulaU = "0"
    '            shp.Cells("FillForegnd").FormulaU = "3"
                
                'Запускаем демона отрисовщика
                cornerDemon.SetStartPosition cell.x, cell.y
                cornerDemon.SetDirection s
                cornerDemon.RunDemon
    
                'Повторяем попытку найти путь - для корректной обработки ситуаций, когда демон может двигаться в двух направлениях (актуально для наклонных стен)
                cornerDemon.SetDirection s
                cornerDemon.RunDemon
            End If
        Next cell

    '    Debug.Print "Точек под заливку" & cornerDemon.lineFillDemon.cellsForFill.Count
        For Each cell In cornerDemon.lineFillDemon.cellsForFill
            RizeCellFireToMax cell.x, cell.y
        Next cell
    End If
    
    
    'Проверяем клетки, на предмет того, не стали ли они внтуренними, и в этом случае удаляем их из коллекций
    For Each cell In matrix.activeFiredCells
        If matrix.IsInner(cell.x, cell.y) Then
            RemoveFromActiveCells cell.x, cell.y
            RemoveFromFrontCells cell.x, cell.y
        End If
        If matrix.IsNotPerimeter(cell.x, cell.y) Then
            RemoveFromPerimeterCells cell.x, cell.y
        End If
    Next cell
    
    'Запускаем анализ наличия пожарнных стволов
'    Dim nozzleDemon As c_FireNozzlesDemon
'    Set nozzleDemon = New c_FireNozzlesDemon
'    nozzleDemon.SetGrain = grain
'    Set nozzleDemon.setMatrix = matrix
    nozzleDemon.RunDemon

    'Возможно имеет смысл оставить в качестве опции (или еще как то) безусловный расчет площади тушения, в независимаости от того, поданы стволы или нет - это важно
    'Рассчитываем площади тушения для каждого имеющегося на странице ствола, при условии, что для него укзана точка контакта с фронтом пожара
    FireSquareCalculate nozzleDemon.nozzles




    'Увеличиваем номер расчитанного шага
    calcStep = calcStep + 1
End Sub


Private Sub Attack(ByVal x As Long, ByVal y As Long)
'Атакуем указанной клеткой все соседние в соовтествии с мощностью горения и весами распространения
Dim firePower As Double
Dim delta1Stvol As Date
'GetFirePower
    
    '0 Проверяем находится ли клетка на листе
    If matrix.IsCellInMatrix(x, y) = False Then
        Exit Sub
    End If
    
    '1 получаем мощность горения клетки
    firePower = matrix.GetFirePower(x, y)
    
    '2 проверяем распространяет ли клетка горение и является ли она внешнеий (позже)
    If firePower = 0 Or matrix.IsInner(x, y) Then
        Exit Sub
    End If
    
    '3 Проверяем какой вариант расчета поданных стволов выбран
'Это все следует обрабатывать в построителе площади тушения (Если выбрана опчия ТОЛЬКО ПЕРВЫЙ СТВОЛ, то площадь тушения оприори считается от всего фронта пожара)
'    If F_InsertFire.CB_StvolsOptions.Text = "первого поданного ствола" Then
'        '3.1 Проверяем находится ли текущее расчетное время в диапазоне 10 минут - время подачи 1 ствола (ВП1) и исходя из этого считаем распространение
'        ' Для работы текущее время должно быть больше 10 и меньше дельты подачи первого ствола (с момента начала пожара)
'        delta1Stvol = DateDiff("n", F_InsertFire.VmD_TimeStart, F_InsertFire.VmD_Time1Stvol)
'        If time > 10 And time < delta1Stvol And isOdd(step) Then
'            Exit Sub
'        End If
'    Else
        '3.2 Проверяем четность хода, время прошедшее с начала пожара (более 10 минут) и подачу воды на тушение
        If Not extSquare Is Nothing Then
            Select Case extSquare.IsCellExtinguished(x, y)
                Case Is = WaterExpenseKind.notSet               'Вода не подается или клетка не находится в зоне тушения
                    'Ничего не делаем - идем дальше и строим как ни в чем не бывало
                Case Is = WaterExpenseKind.notSufficient        'Воды подается недостаточно - считаем только каждый второй раз
                    If isOdd(step) And time > 10 Then
                        Exit Sub
                    End If
                Case Is = WaterExpenseKind.sufficient           'Воды поджается достаточно - не рассчитываем далее
                    Exit Sub
            End Select
        End If
'    End If

    
    '4 Атакуем соседние клетки
    'по диагонали
        RizeGettedCellPower x - 1, y - 1, firePower * diag
        RizeGettedCellPower x - 1, y + 1, firePower * diag
        RizeGettedCellPower x + 1, y - 1, firePower * diag
        RizeGettedCellPower x + 1, y + 1, firePower * diag
    'по ортогонали
        RizeGettedCellPower x - 1, y, firePower * orto
        RizeGettedCellPower x + 1, y, firePower * orto
        RizeGettedCellPower x, y - 1, firePower * orto
        RizeGettedCellPower x, y + 1, firePower * orto
    
End Sub

Private Sub RizeGettedCellPower(ByVal x As Long, ByVal y As Long, ByVal value As Double)
'Увеличиваем полученное значение горения
    
    'Проверяем, не выходит ли координата ячейки за пределы матрицы
    If x < 0 Or y < 0 Or x > matrix.DimensionX Or y > matrix.DimensionY Then Exit Sub
    
    'Проверяем, не входит ли данная клетка в область расчета
    If matrix.IsCellOutOfArea(x, y) Then Exit Sub
    
    'Проверяем, не является ли указанная ячейка стеной
    If matrix.GetOpenSpaceLayerValue(x, y) = 1 Then Exit Sub
    
    'Если до этого клетка не была заряжена, добавляем ее в коллекцию клеток получивших заряд в этом ходу
    If matrix.GetPowerInOneStep(x, y) = 0 Then
        matrix.chargedCells.Add CreateNewCell(x, y)
    End If
    
    'Увеличиваем значение
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, matrix.GetPowerInOneStep(x, y) + value
    
End Sub


Private Sub RizeCellFire(ByVal x As Long, ByVal y As Long)
'Увеличиваем мощность горения клетки
Dim startFirePower As Double
    startFirePower = matrix.GetActualFirePower(x, y)
    
    
    'Если клетка неактивна и мощность полученная ей тоже равна 0, то выходим
    If matrix.GetActualFirePower(x, y) = 0 And matrix.GetPowerInOneStep(x, y) = 0 Then Exit Sub

    'Если клетка изначально была неактивной и при этом загорелась, то первым делом добавляем ее в коллекцию горящих клеток
    If matrix.GetActualFirePower(x, y) = 0 And matrix.GetPowerInOneStep(x, y) > 0 Then
        matrix.firedCells.Add CreateNewCell(x, y)
        'Добавляем клетку в коллекцию активных клеток
        AddToActiveCells x, y
    End If

    'Увеличиваем текущую мощность горения данной клетки
    matrix.RizeCellFire x, y, matrix.GetPowerInOneStep(x, y)
    
    'Если клетка загорелась, но при этом еще горит не на полную мощность, добавляем ее в колекцию загорающихся клеток
    If matrix.GetActualFirePower(x, y) > 0 And matrix.GetActualFirePower(x, y) < matrix.GetMaximumBurnPower Then
        matrix.startedFireCells.Add CreateNewCell(x, y)
    End If
    
    'Добавляем клетку в коллекцию клеток фронта и периметра, если мощность горения клетки максимальна
    If startFirePower < matrix.GetMaximumBurnPower And matrix.GetFirePower(x, y) = matrix.GetMaximumBurnPower Then
        AddToFrontCells x, y
        AddToPerimeterCells x, y
        
        'Добавляем в коллекцию максимально горящих клеток
        matrix.fireMaxCells.Add CreateNewCell(x, y), x & ":" & y
            
        '        'Отрисовываем квадрат
        '        Dim shp As Visio.Shape
        '        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
        '                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
        '        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
        '        shp.Cells("LinePattern").FormulaU = "0"
        '        shp.Cells("FillForegnd").FormulaU = "2"

'            DoEvents
    End If

    
    'Очищаем значение полученной клеткой энергии
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, 0
End Sub

Private Sub RizeCellFireToMax(ByVal x As Long, ByVal y As Long)
'Указываем, что клетка горит на максимум (процедура для дорисовки при переходе к прямоугольной форме)
    'Если клетка изначально была неактивной и при этом загорелась, то первым делом добавляем ее в коллекцию горящих клеток
    If matrix.GetActualFirePower(x, y) = 0 Then
        matrix.firedCells.Add CreateNewCell(x, y)
        'Добавляем клетку в коллекцию активных клеток
'        AddToActiveCells x, y
    End If

    'Увеличиваем текущую мощность горения данной клетки
    matrix.SetFireCell x, y
       
    'Добавляем клетку в коллекцию клеток фронта и периметра, если мощность горения клетки максимальна
'    If startFirePower < matrix.GetMaximumBurnPower And matrix.GetFirePower(x, y) = matrix.GetMaximumBurnPower Then
        
        AddToFrontCells x, y
        AddToPerimeterCells x, y
        AddToActiveCells x, y
            
'        'Отрисовываем квадрат
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'        shp.Cells("LinePattern").FormulaU = "1"
'        shp.Cells("LineColor").FormulaU = "2"
'        shp.Cells("FillPattern").FormulaU = "0"

'            DoEvents
'    End If

    
    'Очищаем значение полученной клеткой энергии
    matrix.SetCellValue x, y, mtrGettedPowerInOneStepLayer, 0
End Sub

Private Sub AddToActiveCells(ByVal x As Long, ByVal y As Long)
'Добавляем клетку в коллекцию активных клеток. Если же клетка с указанным ключем уже имеется - выходим
    On Error Resume Next
    matrix.activeFiredCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromActiveCells(ByVal x As Long, ByVal y As Long)
'Удаляем клетку из коллекции активных клеток. Если же клетка с указанным ключем отсутствует - выходим
    On Error Resume Next
    matrix.activeFiredCells.Remove x & ":" & y
End Sub

Private Sub AddToFrontCells(ByVal x As Long, ByVal y As Long)
'Добавляем клетку в коллекцию клеток фронта. Если же клетка с указанным ключем уже имеется - выходим
    On Error Resume Next
    matrix.fireFrontCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromFrontCells(ByVal x As Long, ByVal y As Long)
'Удаляем клетку из коллекции клеток фронта. Если же клетка с указанным ключем отсутствует - выходим
    On Error Resume Next
    matrix.fireFrontCells.Remove x & ":" & y
End Sub

Private Sub AddToPerimeterCells(ByVal x As Long, ByVal y As Long)
'Добавляем клетку в коллекцию клеток периметра. Если же клетка с указанным ключем уже имеется - выходим
    On Error Resume Next
    matrix.firePerimeterCells.Add CreateNewCell(x, y), x & ":" & y
End Sub
Private Sub RemoveFromPerimeterCells(ByVal x As Long, ByVal y As Long)
'Удаляем клетку из коллекции клеток периметра. Если же клетка с указанным ключем отсутствует - выходим
    On Error Resume Next
    matrix.firePerimeterCells.Remove x & ":" & y
End Sub

Public Sub RefreshFirePerimeter()
'Переводим все клетки находящиеся на границе зоны горения в активное состояние
'это позволяет возобновить расчет при появлении дыр или при изменении конфигурации открытого пространства
Dim cell As c_Cell
    
    'Проверяем клетки, на предмет того, не стали ли они внтуренними, и в этом случае удаляем их из коллекций
    For Each cell In matrix.firedCells
        If Not matrix.IsInner(cell.x, cell.y) Then
            AddToActiveCells cell.x, cell.y
            AddToFrontCells cell.x, cell.y
        End If
    Next cell
End Sub


'-----------------------------Свойства---------------------------------------------
Public Property Get CurrentStep()
'Текущий шаг расчета
    CurrentStep = step
End Property
Public Property Get CalculatedStep()
'рассчитанный шаг расчета
    CalculatedStep = calcStep
End Property

Public Sub RizeCurrentStep()
'Увеличиваем текущий шаг расчета
    step = step + 1
    
'    'Расчитываем пройденный путь и затраченное время
'    distance = GetWayLen(calcStep, grain)
''    diffDistance = currentDistance - prevDistance
'
'    time = distance / speed
''    diffTime = currentTime - prevTime
End Sub



'-----------------------------Инструменты-------------------------------------------
Private Function CreateNewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
    On Error GoTo ex

Dim cell As c_Cell
    
    Set cell = New c_Cell
    Set cell.ParentMatrix = matrix
    cell.x = x
    cell.y = y
    
Set CreateNewCell = cell
Exit Function
ex:
     Debug.Print "err"
End Function

'Private Function GetWayLen(ByVal stepsCount As Integer, ByVal grain As Double) As Single
''Функция возвращает пройденный путь в метрах
'    Dim metersInGrain As Double
'    metersInGrain = grain / 1000
'
'    GetWayLen = CalculateWayLen(stepsCount) * metersInGrain
'End Function
'
'Private Function CalculateWayLen(ByVal stepsCount As Integer) As Integer
''Функция возвращает пройденный путь в клетках
'    Dim tmpVal As Integer
'    tmpVal = 0.58 * stepsCount
'    CalculateWayLen = IIf(tmpVal < 0, 0, tmpVal)
'End Function
'
'Private Function IsEven(ByVal number As Integer) As Boolean
''Проверяем, четное ли число
'    IsEven = Int(number / 2) = number / 2
'End Function

'-----------------------------Разное-------------------------------------------
Public Function GetFiredCellsCount() As Long
    GetFiredCellsCount = matrix.firedCells.Count
End Function
Public Function GetActiveCellsCount() As Long
    GetActiveCellsCount = matrix.activeFiredCells.Count
End Function

Public Function GetFireSquare() As Single
'Возвращаем площадь пожара
    GetFireSquare = matrix.fireMaxCells.Count * grain ^ 2 / 1000000
End Function

Public Function GetExtSquare() As Single
'Возвращает площадь тушения
Dim tmpVal As Single
    If Not extSquare Is Nothing Then
        tmpVal = extSquare.GetSquare
    Else
        tmpVal = 0
    End If
    
    If tmpVal = 0 Then
        tmpVal = GetFireSquare
    End If
    GetExtSquare = tmpVal
End Function

Public Function GetOpenSpaceLayer() As Variant
    GetOpenSpaceLayer = matrix.GetOpenSpaceLayer
End Function

'----------------------------Расчет площадей тушения--------------------------
Public Sub FireSquareCalculate(ByRef nozzleCollection As Collection)
'Процедура получения блоков площади тушения
Dim nozzle As c_FireNozzleModel
Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
    
    Set extSquare = New c_ExtinguishingSquare
    extSquare.intenseNeed = intenseNeed
    
    Set frontDemonImpl = New c_FrontDemon
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    frontDemonImpl.CreateMatrix matrix.DimensionX, matrix.DimensionY
    frontDemonImpl.grain = grain
    frontDemonImpl.SetStartState matrix.fireFrontCells
    
    extinguishingSquareDemon.CreateMatrix matrix
    extinguishingSquareDemon.setGrain = grain
    
    
    Dim i As Integer
    i = 0
    For Each nozzle In nozzleCollection
        If Not nozzle.fireContactCell Is Nothing Then
'            Debug.Assert i < 1
            
'            If i > 0 Then frontDemonImpl.DrawBlockExists extSquare, nozzle.fireContactCell
            'Формируем коллекцию фрагментов площади фронта пожара (и блоков тушения на их основе)
            frontDemonImpl.RunDemon extSquare, matrix.fireFrontCells, nozzle
            
'            'Рассчитываем размер каждого из блоков
'            extinguishingSquareDemon.RunDemon extSquare

            i = i + 1
        End If
    Next nozzle
    
    'Рассчитываем размер каждого из блоков
    extinguishingSquareDemon.RunDemon extSquare
    
    Dim blck As c_ExtinguishingBlock
    Dim extSquareMeters As Single
    Dim needExpense As Single
    Dim factExpense As Single
    Dim expKindForAll As String
    For Each blck In extSquare.extinguishingBlocks
        extSquareMeters = extSquareMeters + blck.GetSquare
        needExpense = needExpense + blck.GetExpenceNeed
        factExpense = factExpense + blck.expenseFact
        expKindForAll = expKindForAll & " " & blck.GetWaterExpenseKind
    Next blck
    
'    If extSquare.extinguishingBlocks.Count > 0 Then
'        Set blck = extSquare.extinguishingBlocks(1)
        Debug.Print "Площадь тушения: " & extSquareMeters
        Debug.Print "Требуемый расход : " & needExpense
        Debug.Print "Фактический расход : " & factExpense
        Debug.Print "Достаточность расхода: " & expKindForAll
'    End If
    
    
    
    
    
    
    
End Sub

Public Sub DrawExtSquareByDemon(ByRef fireShp As Visio.Shape)
'Команда демону отрисовать площадь тушения
'Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
'Dim fireConturs As Visio.Selection
Dim extShp As Visio.Shape
Dim ExtSquareFormula As String
Dim squareShapesCol As Collection

'    Set squareShapesCol = New Collection
    
    'Проверяем активен ли построитель площадей тушения
    If extSquare.extinguishingBlocks.Count = 0 Then
        Exit Sub
    End If
    
    ClearLayer "ExtSquare"
    
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon

    extinguishingSquareDemon.setGrain = grain
    Set squareShapesCol = extinguishingSquareDemon.DrawFireSquare(extSquare)
    
    'Составляем стрку формулы для ссылки на размеры полученных площадей тушения
    For Each extShp In squareShapesCol
        If ExtSquareFormula = "" Then
            ExtSquareFormula = ExtSquareFormula & "IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        Else
            ExtSquareFormula = ExtSquareFormula & "+IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        End If
    Next extShp
    
    'Присваиваем значению площади тушения рассматриваемой фигуры площади пожара значение сумм площадей тушения полченных фигур (формулу)
    fireShp.Cells("Prop.ExtFull").FormulaU = "Index(1, Prop.ExtFull.Format)"
    fireShp.Cells("Prop.ExtSquareT").FormulaU = ExtSquareFormula
End Sub

Public Function GetWaterExpenseKind() As WaterExpenseKind
'Проверяем достаточность расхода воды для тушения
Dim tmpVal As WaterExpenseKind
    
    If Not extSquare Is Nothing Then
        tmpVal = extSquare.GetWaterExpenseKind
    Else
        tmpVal = WaterExpenseKind.notSufficient
    End If
    
GetWaterExpenseKind = tmpVal
End Function

''-------Тестовая процедура
'Public Sub DrawActiveCells()
''Отрисовываем квадрат
'Dim cell As c_Cell
'
'    For Each cell In matrix.activeFiredCells
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Огонь")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'    Next cell
'End Sub
'
'Public Sub DrawFrontCells()
''Отрисовываем квадрат
'Dim cell As c_Cell
'
'    For Each cell In matrix.fireFrontCells
'        Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Огонь")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'    Next cell
'End Sub
'
'Public Sub RemoveActive()
'Dim cell As c_Cell
'
'    For Each cell In matrix.activeFiredCells
'        If matrix.IsInner(cell.x, cell.y) Then
'            RemoveFromActiveCells cell.x, cell.y
'            RemoveFromFrontCells cell.x, cell.y
'        End If
'    Next cell
'End Sub

Public Sub DrawPerimeterCells()
'Отрисовываем фигуру периметра при помощи демона
Dim perimterDemon As c_PerimeterDemon
Dim fireConturs As Visio.Selection
Dim shp As Visio.Shape
    
    ClearLayer "Fire"
    
    Set perimterDemon = New c_PerimeterDemon
    perimterDemon.setGrain(grain).RunDemon matrix.firePerimeterCells
    perimterDemon.DrawSquare
    
    'Объединяем полученные фигуры в окончательную на текущем шаге
'    On Error Resume Next
    Set fireConturs = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    If fireConturs.Count > 0 Then
        fireConturs.Combine
        Set shp = Application.ActiveWindow.Selection(1)
        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
    End If
    
End Sub
