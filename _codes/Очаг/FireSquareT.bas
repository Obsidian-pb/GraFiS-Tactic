Attribute VB_Name = "FireSquareT"
Public fireModeller As c_Modeller
Dim frmF_InsertFire As F_InsertFire
Public grain As Integer

Public stopModellingFlag As Boolean      'Флаг остановки моделирования

'------------------------Модуль для построения площади пожара с использованием тактического метода-------------------------------------------------

Public Sub MakeMatrix(ByRef controlForm As Object)
'Формируем матрицу
Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder
    

    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    

    
    'Запекаем матрицу открытых пространств
    Set matrixBuilder = New c_MatrixBuilder
    matrixBuilder.SetForm controlForm
    matrix = matrixBuilder.NewMatrix(grain)

    'Активируем объект матрицы
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

    'Активируем модельера
    Set fireModeller = New c_Modeller
    fireModeller.SetMatrix matrixObj
    
    'Указываем модельеру значение зерна
    fireModeller.grain = grain

    'Ищем фигуры очага и по их координатам устанавливаем точки начала пожара
    GetFirePoints
    
    controlForm.lblMatrixIsBaked.Caption = "Матрица запечена за " & tmr.GetElapsedTime & " сек."
    controlForm.lblMatrixIsBaked.ForeColor = vbGreen
    
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub

Public Sub RefreshOpenSpacesMatrix(ByRef controlForm As Object)
'Обновляем матрицу открытых пространств
Dim matrix() As Variant
Dim matrixBuilder As c_MatrixBuilder
    
    If fireModeller Is Nothing Then
        MsgBox "Вы не можете обновить не запеченную матрицу!", vbCritical
        Exit Sub
    End If
    
    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    
    'Запекаем матрицу открытых пространств
    Set matrixBuilder = New c_MatrixBuilder
    matrixBuilder.SetForm controlForm
    matrix = matrixBuilder.NewMatrix(grain)
    
    'Обновляем матрицу открытых пространств
    fireModeller.refreshOpenSpaces matrix
    
    'Обновляем периметр пожара
    fireModeller.RefreshFirePerimeter
    
    'Выводим сообщение о итогах обновления
    controlForm.lblMatrixIsBaked.Caption = "Матрица обновлена за " & tmr.GetElapsedTime & " сек."
    controlForm.lblMatrixIsBaked.ForeColor = vbGreen

    tmr.PrintElapsedTime
    Set tmr = Nothing
    
End Sub


Public Sub RunFire(ByVal timeElapsed As Single, ByVal speed As Single, ByVal intenseNeed As Single, Optional ByVal path As Single)
'Моделируем площадь горения до тех пор, пока расчетный путь пройденный огнем не станет больше distance + пройденный ранее (хранится в модельере)
Dim vsO_FireShape As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim newFireShape As Visio.Shape
Dim modelledFireShape As Visio.Shape
Dim borderShape As Visio.Shape

    'Включаем обработчик ошибок - для предупреждения об отсутствии запеченной матрицы
    On Error GoTo ex
    
    'Если путь равен 0, то указываем его бесконечно большим
    If path = 0 Then path = 10000
    
    '---Подключаем таймер
    Dim tmr As c_Timer, tmr2 As c_Timer
    Set tmr = New c_Timer
    Set tmr2 = New c_Timer
    
    Dim i As Integer
    i = 1
    
    '---Определяем предельное значение пройденного пути (путь данного этапа + путь пройденный ранее)
    Dim boundDistance As Single             'Предельное расстояние, согласно расчета
    Dim currentDistance As Single           'Текущее пройденное расстояние
    Dim prevDistance As Single              'Расстояние пройденное на предыдущем этапе расчета
    Dim diffDistance As Single              'Расстяоине пройденное в данном этапе расчета
    Dim realCurrentDistance As Single       'Реальное Текущее пройденное расстояние
    Dim realDiffDistance As Single          'Реальное Расстяоине пройденное в данном этапе расчета
    Dim currentTime As Single               'Текущее время с начала расчета
    Dim prevTime As Single                  'Время за которое проейден предыдущий этап расчета
    Dim diffTime As Single                  'Время за которое проейден текущий этап расчета
    
    '---Activate nozzles calculations
    fireModeller.ActivateNozzles F_InsertFire
    
    'Указываем модельеру значение требуемой интенсивности подачи воды
    fireModeller.intenseNeed = intenseNeed
    
    
    prevDistance = fireModeller.distance
    boundDistance = timeElapsed * speed + prevDistance
    
    prevTime = fireModeller.time
    
    Do While diffTime < timeElapsed And realCurrentDistance < path
        ClearLayer "ExtSquare"
        
'        Stop   ' - Здесь нужно добавить проверку на достаточность расхода для тушения -> fireModeller.GetExtSquare
        If fireModeller.GetExtSquare < fireModeller.GetFireSquare Then
'        fireModeller.
            'Проверяем, сколько времени длится расчет, если меньше 10 минут, то расчитываем, только каждый второй шаг, т.е., с половиной скорости
            If currentTime < 10 Then
                'При вермени менее 10 минут считаем рост только каждый второй шаг
                If IsEven(fireModeller.CurrentStep) Then
                    fireModeller.OneRound
    
                    'Объединяем добавленные точки в одну фигуру
                    MakeShape
                End If
            Else
                fireModeller.OneRound
                    
                'Объединяем добавленные точки в одну фигуру
                MakeShape
            End If
        ElseIf fireModeller.GetExtSquare >= fireModeller.GetFireSquare Then
            If Not fireModeller.GetWaterExpenseKind = sufficient Then   'Если достаточно расхода то ничего не делаем, просто считаем следующий шаг
            
'            Else
                'Проверяем, сколько времени длится расчет, если меньше 10 минут, то расчитываем, только каждый второй шаг, т.е., с половиной скорости
                If currentTime < 10 Then
                    'При вермени менее 10 минут считаем рост только каждый второй шаг
                    If IsEven(fireModeller.CurrentStep) Then
                        fireModeller.OneRound
        
                        'Объединяем добавленные точки в одну фигуру
                        MakeShape
                    End If
                Else
                    fireModeller.OneRound
                        
                    'Объединяем добавленные точки в одну фигуру
                    MakeShape
                End If
            End If
        End If
        
        'Увеличиваем шаг расчета
        fireModeller.RizeCurrentStep
            
        currentDistance = GetWayLen(fireModeller.CurrentStep, grain)
        diffDistance = currentDistance - prevDistance
        realCurrentDistance = GetWayLen(fireModeller.CalculatedStep, grain)
        realDiffDistance = realCurrentDistance - prevDistance
        
        currentTime = currentDistance / speed
        diffTime = currentTime - prevTime
               
        On Error Resume Next
        '---Печатаем сколько потребовалось времени
        F_InsertFire.lblCurrentStatus.Caption = "Шаг: " & i & "(" & fireModeller.CurrentStep & "), " & _
                                                " пройденный путь: " & Round(realDiffDistance, 2) & "(" & Round(realCurrentDistance, 2) & ")м.," & _
                                                " время: " & Round(diffTime, 2) & "(" & Round(currentTime, 2) & ")мин, " & _
                                                Chr(13) & "Площадь пожара: " & fireModeller.GetFireSquare & "м.кв., " & _
                                                Chr(13) & "Площадь тушения: " & fireModeller.GetExtSquare & "м.кв., " & _
                                                Chr(13) & "Требуемый расход: " & fireModeller.GetExtSquare * fireModeller.intenseNeed & "л/с"
        'Указываем форме настроек время прошедшее с начала моделирования
        F_InsertFire.timeElapsedMain = currentTime
        'Указываем форме настроек путь пройденный с начала моделирования
        F_InsertFire.pathMain = realCurrentDistance
        
        
        On Error GoTo ex
        
        i = i + 1
        
        fireModeller.distance = realCurrentDistance ' currentDistance
        fireModeller.time = currentTime
               
        'Очищаем выделение и выполняем команды пользователя
        Application.ActiveWindow.DeselectAll
        DoEvents
        
        'Если пользователь нажал в форме кнопку "Остановить" прекращаем моделирвоание
        If stopModellingFlag Then
            Exit Do
        End If
    Loop
        
    '---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    Set modelledFireShape = vsoSelection(1)
    Application.ActiveWindow.Select modelledFireShape, visSelect
    
    '---Собственно обращение
    ImportAreaInformation
'    '---Указываем для фигуры фактическую площадь тушения
    If fireModeller.GetExtSquare > 0 And F_InsertFire.flag_DrawExtSquare.value = True Then
        fireModeller.DrawExtSquareByDemon modelledFireShape
    End If
    'Перемещаем полученные фигуры на задний план
    modelledFireShape.SendToBack
    
    'Перемещаем фигуру расчетной зоны (при ее наличии) на задний план
    If TryGetShape(borderShape, "User.IndexPers:1001") Then
        borderShape.SendToBack
    End If
        
''TEST:
'fireModeller.DrawExtSquareByDemon modelledFireShape
'Ставим фокус на построенной ранее фигуре зоны горения
Application.ActiveWindow.DeselectAll
Application.ActiveWindow.Select modelledFireShape, visSelect

        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
    
Exit Sub
ex:
    MsgBox "Матрица не запечена!", vbCritical
    
    '---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
    Set modelledFireShape = vsoSelection(1)
    Set newFireShape = ActivePage.Drop(modelledFireShape, _
                        modelledFireShape.Cells("PinX").Result(visInches), modelledFireShape.Cells("PinY").Result(visInches))
    
    '---Собственно обращение
    ImportAreaInformation
    'Перемещаем полученные фигуры на задний план
    newFireShape.SendToBack
    modelledFireShape.SendToBack
        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
End Sub

' уничтожение матрицы (очищаем память)
Public Sub DestroyMatrix()
    Set fireModeller = Nothing
End Sub

Public Function IsAcceptableMatrixSize(ByVal maxMatrixSize As Long, ByVal grain As Integer) As Boolean
Dim xCount As Long
Dim yCount As Long
Dim shp As Visio.Shape

    On Error GoTo ex
    
    'Проверяем нет ли на данной страницы фигуры расчетной зоны. Если есть, то определяем, что расчет возможен
    If TryGetShape(shp, "User.IndexPers:1001") Then
        IsAcceptableMatrixSize = True
        Exit Function
    End If
'    grain = Me.txtGrainSize.value

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
    
    IsAcceptableMatrixSize = xCount * yCount < maxMatrixSize
Exit Function
ex:
    IsAcceptableMatrixSize = False
End Function




'Не понял откуда это
'Public Sub DrawExtSquare()
''Внешняя команда на отрисовку площади тушения
'    fireModeller.DrawExtSquareByDemon
'End Sub









Private Sub GetFirePoints()
'Модуль ищет и указывает точки начала горения
Dim shp As Visio.Shape

    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            If shp.Cells("User.IndexPers") = 70 Then
                '---Устанваливаем старотовую точку, для дальнейшего расчета распространения огня
                SetFirePointFromCoordinates shp.Cells("PinX").Result(visMillimeters), _
                    shp.Cells("PinY").Result(visMillimeters)
            End If
        End If
    Next shp
   
End Sub

Private Sub SetFirePointFromCoordinates(xPos As Double, yPos As Double)
'Отмечаем в матрице горящую клетку по пришедшим геометрическим координатам
Dim xIndex As Integer
Dim yIndex As Integer

    xIndex = Int(xPos / grain)
    yIndex = Int(yPos / grain)
    
    fireModeller.SetStartFireCell xIndex, yIndex

End Sub

Private Sub MakeShape()
'Отрисовываем фигуру хоны горения при помощи демона
    fireModeller.DrawPerimeterCells
End Sub

Public Function GetStepsCount(ByVal grain As Integer, ByVal speed As Single, ByVal elapsedTime As Single) As Integer
'Функция возвращает количество шагов в зависиомсти от размера зерна, скорости распространения огня и времени на которое производится расчет

    '1 определить путь который должен пройти огонь
    Dim firePathLen As Double
    firePathLen = speed * elapsedTime * 1000 / grain
    
    '2 определить собственно сколько нужно шагов для достижения
    Dim tmpVal As Integer
    tmpVal = firePathLen / 0.58

    GetStepsCount = IIf(tmpVal < 0, 0, tmpVal)
    
End Function

Public Function GetWayLen(ByVal stepsCount As Integer, ByVal grain As Double) As Single
'Функция возвращает пройденный путь в метрах
    Dim metersInGrain As Double
    metersInGrain = grain / 1000

    GetWayLen = CalculateWayLen(stepsCount) * metersInGrain
End Function

Public Function CalculateWayLen(ByVal stepsCount As Integer) As Integer
'Функция возвращает пройденный путь в клетках
    Dim tmpVal As Integer
    tmpVal = 0.58 * stepsCount
    CalculateWayLen = IIf(tmpVal < 0, 0, tmpVal)
End Function




Public Function IsMatrixBacked() As Boolean
'Возвращает Истина, если матрица уже запечена и Ложь, если нет
    IsMatrixBacked = Not fireModeller Is Nothing
End Function

Private Function IsEven(ByVal number As Integer) As Boolean
'Проверяем, четное ли число
    IsEven = Int(number / 2) = number / 2
End Function


'------------------------------------Добавление к площади пожара указанной фигуры------------------------------
Public Sub AddFireArea(ShpObj As Visio.Shape)
'Добавление к площади пожара указанной фигуры
        
    If Not IsMatrixBacked Then
        MsgBox "Матрица не запечена!!!"
        Exit Sub
    End If
    
    'Анализируем фигуру и добавляем зону горения в расчет
    fireModeller.AddFireFromShape ShpObj

    MsgBox "Площадь фигуры добавлена к площади горения." & Chr(13) & Chr(13) & _
            "ОБРАТИТЕ ВНИМАНИЕ, что выбранная фигура сохранится на листе и будет учитыватья при анализе! Чтобы избежать этого, удалите ее!"
End Sub
