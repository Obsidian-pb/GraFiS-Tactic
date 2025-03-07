Attribute VB_Name = "FireSquareModel"
Public fireModeller As c_Modeller
Dim frmF_InsertFire As F_InsertFire
'Public GRAIN As Integer

Public stopModellingFlag As Boolean      'Флаг остановки моделирования
Dim roundNumber As Integer

'Public Const FIRESPEED = 1
Public Const FIREINTENSE = 0.1
'Public Const GRAIN = 200


Dim exl As Object
Dim wkb As Object
Dim sht As Object
Dim erow As Integer


'------------------------Модуль для построения площади пожара с использованием тактического метода-------------------------------------------------

Public Sub MakeMatrixM(ByRef controlForm As Object, ByVal roundsCount As Integer, ByVal calcTime As Integer, ByVal checkTime As Integer)
'Формируем матрицу
Dim matrix() As Variant
Dim matrixObj As c_Matrix
Dim matrixBuilder As c_MatrixBuilder

Dim linespeed As Single
    
    On Error GoTo EX1
    linespeed = CStr(controlForm.tb_speed.value)


    On Error GoTo EX2
    '---Создаем новый документ Эксель
    Set exl = CreateObject("Excel.Application")
    exl.Visible = True
    Set wkb = exl.Workbooks.Add
    Set sht = wkb.Sheets(1)

    '---Подключаем таймер
    Dim tmr As c_Timer
    

    Set tmr = New c_Timer
    
    
    'Запекаем матрицу открытых пространств
    Set matrixBuilder = New c_MatrixBuilder
    controlForm.SetMatrixSize 0
    matrixBuilder.SetForm controlForm
    matrix = matrixBuilder.NewMatrix(GRAIN)

    'Активируем объект матрицы
    Set matrixObj = New c_Matrix
    matrixObj.CreateMatrix UBound(matrix, 1), UBound(matrix, 2)
    matrixObj.SetOpenSpace matrix
        

    controlForm.lblMatrixIsBaked.Caption = "Матрица запечена за " & tmr.GetElapsedTime & " сек."
    controlForm.lblMatrixIsBaked.ForeColor = vbGreen
    Set tmr = Nothing
    
    For roundNumber = 1 To roundsCount
        erow = roundNumber
        
        'Активируем модельера
        Set fireModeller = New c_Modeller
        fireModeller.SetMatrix matrixObj
        
        'Указываем модельеру значение зерна
        fireModeller.GRAIN = GRAIN
    
        'Ищем фигуры очага и по их координатам устанавливаем точки начала пожара
        GetFirePoints matrixObj
        
        'Запускаем моделирование пожара
        stopModellingFlag = False
        RunFireM controlForm, calcTime, linespeed, FIREINTENSE, , checkTime
        
        'Сбрасываем состояние матрицы
        matrixObj.DropState
    Next roundNumber
    
    Debug.Print "Ok"

Exit Sub
EX1:
    MsgBox "Не верно указан формат числа. Попробуйте заменить '.' на ',' или наоборот, в зависиомти от натсроек системы."
EX2:
    MsgBox "Ошибка."
End Sub


Public Sub RunFireM(ByRef controlForm As Object, ByVal timeElapsed As Single, ByVal speed As Single, ByVal intenseNeed As Single, Optional ByVal path As Single, Optional ByVal checkTime As Single)
'Моделируем площадь горения до тех пор, пока расчетный путь пройденный огнем не станет больше distance + пройденный ранее (хранится в модельере)
Dim vsO_FireShape As Visio.Shape
Dim vsoSelection As Visio.Selection
Dim newFireShape As Visio.Shape
Dim modelledFireShape As Visio.Shape
Dim borderShape As Visio.Shape

'Dim erow As Integer
Dim ecol As Integer

    ecol = 1

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
    Dim tickTime As Integer                 'Время вывода результата
    tickTime = checkTime
    
    '---Activate nozzles calculations
    fireModeller.ActivateNozzles F_InsertFire
    
    'Указываем модельеру значение требуемой интенсивности подачи воды
    fireModeller.intenseNeed = intenseNeed
    
    
    prevDistance = fireModeller.distance
    boundDistance = timeElapsed * speed + prevDistance
    
    prevTime = fireModeller.time
    
    'Обновляем расходы на тушение - нжно, для того, чтобы при удалении и перемещении стволов возобновлялся рост площади
    fireModeller.NozllesRecalculate
    
    Do
        ClearLayer "ExtSquare"
        
'        Stop   ' - Здесь нужно добавить проверку на достаточность расхода для тушения -> fireModeller.GetExtSquare
        
        'Если размер площади тушения меньше площади пожара:
'        If fireModeller.GetExtSquare < fireModeller.GetFireSquare Then
            'Проверяем, сколько времени длится расчет, если меньше 10 минут, то расчитываем, только каждый второй шаг, т.е., с половиной скорости
            If currentTime < 10 Then
                'При вермени менее 10 минут считаем рост только каждый второй шаг
                If IsEven(fireModeller.CurrentStep) Then
                    fireModeller.OneRound
    
                    'Объединяем добавленные точки в одну фигуру
                    If controlForm.cb_visualize.value Then
                        MakeShape
                    End If
                End If
            Else
                fireModeller.OneRound
                    
                'Объединяем добавленные точки в одну фигуру
                If controlForm.cb_visualize.value Then
                    MakeShape
                End If
            End If
'        ElseIf fireModeller.GetExtSquare >= fireModeller.GetFireSquare Then
'            If Not fireModeller.GetWaterExpenseKind = sufficient Then   'Если достаточно расхода то ничего не делаем, просто считаем следующий шаг
''                MakeShape
''            Else
'                'Проверяем, сколько времени длится расчет, если меньше 10 минут, то расчитываем, только каждый второй шаг, т.е., с половиной скорости
'                If currentTime < 10 Then
'                    'При вермени менее 10 минут считаем рост только каждый второй шаг
'                    If IsEven(fireModeller.CurrentStep) Then
'                        fireModeller.OneRound
'
'                        'Объединяем добавленные точки в одну фигуру
'                        MakeShape
'                    End If
'                Else
'                    fireModeller.OneRound
'
'                    'Объединяем добавленные точки в одну фигуру
'                    MakeShape
'                End If
'            End If
'        End If
        
        'Увеличиваем шаг расчета
        fireModeller.RizeCurrentStep
            
                    'Возможно это стоит вынести в сам модельер
        currentDistance = GetWayLen(fireModeller.CurrentStep, GRAIN)
        diffDistance = currentDistance - prevDistance
        realCurrentDistance = GetWayLen(fireModeller.CalculatedStep, GRAIN)
        realDiffDistance = realCurrentDistance - prevDistance
        
        currentTime = currentDistance / speed
        diffTime = currentTime - prevTime
               
        On Error Resume Next
        '---Печатаем сколько потребовалось времени
'        Debug.Print "Шаг: " & i & "(" & fireModeller.CurrentStep & "), " & _
'                                                " пройденный путь: " & Round(realDiffDistance, 2) & "(" & Round(realCurrentDistance, 2) & ")м.," & _
'                                                " время: " & Round(diffTime, 2) & "(" & Round(currentTime, 2) & ")мин, " & _
'                                                Chr(13) & "Площадь пожара: " & fireModeller.GetFireSquare & "м.кв., " & _
'                                                Chr(13) & "Площадь тушения: " & fireModeller.GetExtSquare & "м.кв., " & _
'                                                Chr(13) & "Требуемый расход: " & fireModeller.GetExtSquare * fireModeller.intenseNeed & "л/с"
        
        controlForm.lblCurrentStatus.Caption = "Расчет №" & str(roundNumber) & _
                                            Chr(13) & "Шаг: " & i & " время: " & Round(currentTime, 2) & "мин, " & _
                                            Chr(13) & "Площадь пожара: " & fireModeller.GetFireSquare & "м.кв., "
        If currentTime >= tickTime Then
'            Debug.Print "Шаг: " & i & " время: " & Round(currentTime, 2) & "мин, " & _
'                                            "Площадь пожара: " & fireModeller.GetFireSquare & "м.кв., "
            'отправляем в эксель
            sht.Cells(erow, ecol).Formula = fireModeller.GetFireSquare
            ecol = ecol + 1
            tickTime = tickTime + checkTime
        End If
'        'Указываем форме настроек время прошедшее с начала моделирования
'        F_InsertFire.timeElapsedMain = currentTime
'        'Указываем форме настроек путь пройденный с начала моделирования
'        F_InsertFire.pathMain = realCurrentDistance
        
        
        On Error GoTo EX
        
        i = i + 1
        
        fireModeller.distance = realCurrentDistance ' currentDistance
        fireModeller.time = currentTime
               
        'Очищаем выделение и выполняем команды пользователя
        Application.ActiveWindow.DeselectAll
        DoEvents
        
        'Если достигнуты пределы моделирвоания, выходим из цикла
'        If F_InsertFire.OB_ByRadius = True Then
'            If realCurrentDistance >= path Then
'                stopModellingFlag = True
'            End If
'        Else
        If diffTime >= timeElapsed Or realCurrentDistance >= path Then
            stopModellingFlag = True
        End If
'        End If

        
        'Если пользователь нажал в форме кнопку "Остановить" прекращаем моделирвоание
        If stopModellingFlag Then
            Exit Do
        End If
    Loop
        
'    '---Если происходит тушение и расхода достаточно, то значит ранее вигура построена не была и ее необходимо построить
'    If fireModeller.GetExtSquare >= fireModeller.GetFireSquare Then
'        If fireModeller.GetWaterExpenseKind = sufficient Then   'Если достаточно расхода то ничего не делаем, просто считаем следующий шаг
'            MakeShape
'        End If
'    End If
    
    '---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
    If controlForm.cb_visualize.value Then
        Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
        Set modelledFireShape = vsoSelection(1)
        Application.ActiveWindow.Select modelledFireShape, visSelect
        Application.ActiveWindow.Selection.Delete
    End If
    
'    '---Собственно обращение
'    ImportAreaInformation
''    '---Указываем для фигуры фактическую площадь тушения
'    If fireModeller.GetExtSquare > 0 And F_InsertFire.flag_DrawExtSquare.value = True Then
'        fireModeller.DrawExtSquareByDemon modelledFireShape
'    End If
'    'Перемещаем полученные фигуры на задний план
'    modelledFireShape.SendToBack
'
'    'Перемещаем фигуру расчетной зоны (при ее наличии) на задний план
'    If TryGetShape(borderShape, "User.IndexPers:1001") Then
'        borderShape.SendToBack
'    End If
'
'    'Построение фигуры площади тушения по итогам сессии моделирвоания
'    If F_InsertFire.flag_DrawExtSquare Then
'        fireModeller.DrawExtSquareByDemon modelledFireShape
'    End If
'
'    'Ставим фокус на построенной ранее фигуре зоны горения
'    Application.ActiveWindow.DeselectAll
'    Application.ActiveWindow.Select modelledFireShape, visSelect

        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
    
Exit Sub
EX:
    MsgBox "Что-то пощло не так! Результаты данного моделирвоания лучше не учитывать", vbCritical
    
'    '---Определяем получившуюся фигуру и обращаем ее в фигуру площади горения
'    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Fire")
'    Set modelledFireShape = vsoSelection(1)
'    Set newFireShape = ActivePage.Drop(modelledFireShape, _
'                        modelledFireShape.Cells("PinX").Result(visInches), modelledFireShape.Cells("PinY").Result(visInches))
'
'    '---Собственно обращение
'    ImportAreaInformation
'    'Перемещаем полученные фигуры на задний план
'    newFireShape.SendToBack
'    modelledFireShape.SendToBack
        
    Debug.Print "Всего затрачено " & tmr2.GetElapsedTime & "с."
    
    Set tmr = Nothing
    Set tmr2 = Nothing
End Sub

Private Sub MakeShape()
'Отрисовываем фигуру хоны горения при помощи демона
    fireModeller.DrawPerimeterCells
End Sub



















Private Sub GetFirePoints(ByRef matrix As c_Matrix)
'Модуль ищет и указывает точки начала горения
Dim shp As Visio.Shape
Dim i As Integer
Dim X As Long
Dim Y As Long

    i = 0
    Do
        
        X = Int(Rnd() * matrix.DimensionX)
        Y = Int(Rnd() * matrix.DimensionY)
        
        Debug.Print str(i) & ": " & str(X) & ", " & str(Y)
        If matrix.IsCellCanFire(X, Y) Then
            fireModeller.SetStartFireCell X, Y
            Exit Sub
        End If
        
        i = i + 1
        If i > 100 Then
            Exit Do
        End If
    Loop
    
   
End Sub












'---------------------Экспорт в Excel---------------------------------
Public Sub ExportToExcel()
'Экспортируем содержимое списка в документ Excel
Dim i As Integer
Dim j As Integer
Dim colCount As Byte

Dim s As String
        
    
    'Заполняем таблицу ChrW(9500)
    i = 0
    Do Until IsNull(Me.LB_List.Column(1, i))
        For j = 1 To colCount
            s = Me.LB_List.Column(j, i)
            s = Replace(s, ChrW(9500), "")
            s = Replace(s, ChrW(9492), "")
            sht.Cells(i + 1, j).Formula = s
        Next j
        i = i + 1
        If i > 2000 Then
            'аварийный выход
            Exit Do
        End If
    Loop
    
'    exl.Selection.Columns.AutoFit    'Устанавливаем ширину столбцов по содержимому
End Sub


