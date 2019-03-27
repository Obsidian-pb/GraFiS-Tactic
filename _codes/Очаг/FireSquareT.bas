Attribute VB_Name = "FireSquareT"
Dim fireModeller As c_Modeller
Dim frmF_InsertFire As F_InsertFire
Public grain As Integer

Public stopModellingFlag As Boolean      'Флаг остановки моделирования

'------------------------Модуль для построения площади пожара с использованием тактического метода-------------------------------------------------

'Public Sub ShowModellerF_InsertFire()
''    Set frmF_InsertFire = New F_InsertFire
'    F_InsertFire.Show
'End Sub



Public Sub MakeMatrix()

Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder
    

    '---Подключаем таймер
    Dim tmr As c_Timer
    Set tmr = New c_Timer
    

    
    'Запекаем матрицу открытых пространств
    Set matrixBuilder = New c_MatrixBuilder
    matrixBuilder.SetForm F_InsertFire
    matrix = matrixBuilder.NewMatrix(grain)

    'Активируем объект матрицы
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix

    'Активируем модельера
    Set fireModeller = New c_Modeller
    fireModeller.setMatrix matrixObj
    
    'Указываем модельеру значение зерна
    fireModeller.grain = grain
    
'    'Указываем модельеру значение требуемой интенсивности подачи воды
'    fireModeller.intenseNeed = 0.1          'ВРЕМЕННО 0,1 - потом нужно сделать указание из формы!!!

    'Ищем фигуры очага и по их координатам устанавливаем точки начала пожара
    GetFirePoints

    '---Печатаем сколько потребовалось времени
'    MsgBox "Матрица запечена за " & tmr.GetElapsedTime & " сек." & Chr(10) & Chr(13) & "Зерно " & grain & "мм."
    
    F_InsertFire.lblMatrixIsBaked.Caption = "Матрица запечена за " & tmr.GetElapsedTime & " сек."
    F_InsertFire.lblMatrixIsBaked.ForeColor = vbGreen
    
    Debug.Print "Матрица запечена..."
    tmr.PrintElapsedTime
    Set tmr = Nothing


End Sub


Public Sub RunFire(ByVal timeElapsed As Single, ByVal speed As Single, ByVal intenseNeed As Single, Optional ByVal path As Single)
'Моделируем площадь горения до тех пор, пока расчетный путь пройденный огнем не станет больше distance + пройденный ранее (хранится в модельере)
Dim vsO_FireShape As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim newFireShape As Visio.Shape
Dim modelledFireShape As Visio.Shape

    'Включаем обработчик ошибок - для предупреждения об отсутствии запеченной матрицы
    On Error GoTo EX
    
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
    fireModeller.ActivateNozzles
    
    'Указываем модельеру значение требуемой интенсивности подачи воды
    fireModeller.intenseNeed = intenseNeed
    
    
    prevDistance = fireModeller.distance
    boundDistance = timeElapsed * speed + prevDistance
    
    prevTime = fireModeller.time
    
    Do While diffTime < timeElapsed And realCurrentDistance < path
        ClearLayer "ExtSquare"
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
                                                " время: " & Round(diffTime, 2) & "(" & Round(currentTime, 2) & ")мин " & _
                                                "Площадь пожара: " & fireModeller.GetFireSquare & "м.кв." ' & _
                                                "Площадь тушения: " & fireModeller.GetExtSquare & "м.кв."
        'Указываем форме настроек время прошедшее с начала моделирования
        F_InsertFire.timeElapsedMain = currentTime
        'Указываем форме настроек путь пройденный с начала моделирования
        F_InsertFire.pathMain = realCurrentDistance
        
        
        On Error GoTo EX
        
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
    
'    fireModeller.distance = realCurrentDistance ' currentDistance
'    fireModeller.time = currentTime
        
    '---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
'    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
'    Dim newFireShape As Visio.Shape
'    Dim modelledFireShape As Visio.Shape
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
    
Exit Sub
EX:
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

Public Sub DrawExtSquare()
'Внешняя команда на отрисовку площади тушения
    fireModeller.DrawExtSquareByDemon
End Sub












'Public Sub DrawActive()
'    fireModeller.DrawActiveCells
''    fireModeller.DrawFrontCells
'End Sub
''Public Sub RemoveActive()
''    fireModeller.RemoveActive
'''    fireModeller.DrawFrontCells
''End Sub


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

    On Error Resume Next

    Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")

    vsoSelection.Union

    Application.ActiveWindow.Selection(1).CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Fire")
'    Application.ActiveWindow.Selection(1).SendToBack
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
