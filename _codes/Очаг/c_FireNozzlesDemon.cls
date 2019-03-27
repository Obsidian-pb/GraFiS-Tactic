VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireNozzlesDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public nozzles As Collection                        'Коллекция пожарных стволов размещенных на листе

Dim matrix As Variant                               'Матрица расчетной зоны
Dim calculatedCells As Variant                      'Матрица рассчитанных клеток
Private fireMatrixObject As c_Matrix                 'Объект расчетной зоны (внешний, ссылка)

Private grain As Integer                             'Зерно расчета

Const halfRadian As Double = 1.570796326
Const diag As Double = 1.4142135623731              'модификатор распространения вверх и вниз
Const orto As Double = 1                            'модификатор распротсранения по диагонали

Const mmInInch As Double = 25.4

'------------------Класс демона расчетчика стволов----------------------------
'1 - Находит все стволы расположенные на листе
'2 - Строит зоны досягаемости для стволов (рекурсивно)
'3 - Исходя из точек досягаемости стволов отдает команду на расчет соответсвующих блоков площади тушения



Private Sub Class_Initialize()
    Set nozzles = New Collection
End Sub

Public Property Let SetGrain(ByVal grainValue)
    grain = grainValue
End Property

Public Property Set setMatrix(ByRef matr As c_Matrix)
    Set fireMatrixObject = matr
End Property






Public Sub RunDemon()
'Запускаем демона
Dim nozzle As c_FireNozzleModel
    
'    'Формируем коллекцию расположенных на странице стволов
'    FindNozzles
    
    'Перебираем все стволы и проверяем, касается ли их зона тушения (5м для ручных и 10 для лафетных) фронта пожара
    For Each nozzle In nozzles
        ReDim matrix(fireMatrixObject.DimensionX, fireMatrixObject.DimensionY)
        ReDim calculatedCells(fireMatrixObject.DimensionX, fireMatrixObject.DimensionY)
        
        'Проверяем, находится ли ствол в пределах страницы
'        Debug.Print UBound(calculatedCells, 1) & " " & UBound(calculatedCells, 2)
        If nozzle.position.x < 0 Or nozzle.position.y < 0 Or _
                nozzle.position.x > UBound(calculatedCells, 1) Or nozzle.position.y > UBound(calculatedCells, 2) Then
        Else
            calculatedCells(nozzle.position.x, nozzle.position.y) = 1
            'Если зона досягаемости ствола касается где-либо фронта, далее отдаем команду модельеру на расчет блока тушения
            checkFireIntercourse nozzle
        End If
        

    Next nozzle
    
    
    

End Sub




Public Sub FindNozzles()
'Формируем коллекцию стволов
Dim shp As Visio.Shape
Dim nozzle As c_FireNozzleModel
Dim indexPers As Integer
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.IndexPers", 0) Then
            indexPers = shp.Cells("User.IndexPers").Result(visNumber)
            If indexPers = 34 Or indexPers = 36 Or indexPers = 39 Then      '34 - ручной водяной ствол, 36 - лафетный водяной, 39 - возимый водяной
                'Проверка на маневренность
                If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
                    Set nozzle = New c_FireNozzleModel
                    If indexPers = 34 Then      '34 - ручной водяной ствол, 36 - лафетный водяной, 39 - возимый водяной
                        nozzle.nozzleType = waterHand
                    ElseIf indexPers = 36 Or indexPers = 39 Then
                        nozzle.nozzleType = waterLafet
                    End If
                    Set nozzle.position = GetNozzlePosition(shp.Cells("PinX").Result(visMillimeters), _
                        shp.Cells("PinY").Result(visMillimeters))
                    nozzle.waterExpense = shp.Cells("User.PodOut").Result(visNumber)
                    nozzle.angle = shp.Cells("Angle").Result(visNumber) + halfRadian
                    nozzles.Add nozzle
                End If
            End If
        End If
    Next shp
End Sub

Private Function GetNozzlePosition(ByVal xPos As Double, ByVal yPos As Double) As c_Cell
'Находим координату расположения ствола
Dim x As Integer
Dim y As Integer

    x = Int(xPos / grain)
    y = Int(yPos / grain)
    
    Set GetNozzlePosition = NewCell(x, y)
End Function




Private Sub checkFireIntercourse(ByRef nozzle As c_FireNozzleModel)
'Проверяем пересечение направления работы ствола с пожаром
Dim coll As Collection
Dim maxWay As Integer

Dim cs As Double
Dim sn As Double
Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim cell As c_Cell

    If F_InsertFire.optValue Then
        On Error GoTo EX
        maxWay = F_InsertFire.txtNozzleRangeValue * 1000
        On Error GoTo EX1
    Else
        If nozzle.nozzleType = waterLafet Then
            maxWay = 10000
        Else
            maxWay = 5000
        End If
    End If
    
'    Set coll = New Collection
'    coll.Add nozzle.position
'
'    Do While coll.Count > 0
'        Set coll = OneStep(nozzle, coll, maxWay)
'        If coll Is Nothing Then Exit Do
'    Loop
    

    'Проверяем пересечение по направлению пересечения со стволом
        'Заранее определяем косинус и синус угла - для дальнейшего построения линии
        cs = Cos(nozzle.angle)
        sn = Sin(nozzle.angle)
        
        'Последовательно проверяем все клеткии линии, на предмет того, пересекаются ли они с фронтом пожара
        i = 0
        Set cell = New c_Cell
        cell.x = nozzle.position.x
        cell.y = nozzle.position.y
        Do While CellType = 0
            cell.x = nozzle.position.x + Int(cs * i)
            cell.y = nozzle.position.y + Int(sn * i)
'------------------
'            Dim shp As Visio.Shape
'        Set shp = Application.ActivePage.DrawRectangle(cell.x * grain / mmInInch, cell.y * grain / mmInInch, _
'                (cell.x + 1) * grain / mmInInch, (cell.y + 1) * grain / mmInInch)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("spray")
'        shp.Cells("LinePattern").FormulaU = "0"
'        shp.Cells("FillForegnd").FormulaU = "2"
'------------------
            
'            Debug.Print cell.x & ", " & cell.y & " -> " & nozzle.position.x & ", " & nozzle.position.y
            
            If IsIntercourseWithFront2(nozzle, cell) Then
                Exit Do
            End If
            If Math.Sqr((cell.x - nozzle.position.x) ^ 2 + (cell.y - nozzle.position.y) ^ 2) * grain > maxWay Then
                Exit Do
            End If
            i = i + 1
        Loop
    
Exit Sub
EX:
    MsgBox "Не правильно указано значение расстояния рассчета стволов"
Exit Sub
EX1:
    
End Sub


Private Function OneStep(ByRef nozzle As c_FireNozzleModel, ByRef inColl As Collection, ByVal maxWay As Integer) As Collection
'Функция возвращает Истину, если клетка является клеткой фронта пожара (имеет рядом хоть одну горящую не полностью клетку)
'На входе получаем коллекцию текущей волны клеток, на выходе следующую после одного шага распространения волны
Dim cell As c_Cell
Dim newColl As Collection
Dim cellWay As Single
    
    Set newColl = New Collection
    
    For Each cell In inColl
'        cellWay = matrix(cell.x, cell.y)
'        'Распространение по прямой
'            'Заранее определяем косинус и синус угла - для дальнейшего построения линии
'            cs = Cos(nozzle.angle)
'            sn = Sin(nozzle.angle)
'
'            'Последовательно проверяем все клеткии линии, на предмет того, пересекаются ли они с фронтом пожара
'            x = cell.x + Int(cs * i)
'            y = cell.y + Int(sn * i)
'            Set cell = NewCell(x, y)
'
'            If IsIntercourseWithFront(nozzle, newColl, cell, _
'                                        Math.Sqr((x - nozzle.position.x) ^ 2 + (y - nozzle.position.y) ^ 2) * grain, maxWay) Then
'                Set OneStep = Nothing
'                Exit Function
'            End If
        
        'Распространяем на соседние клетки
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, l), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, lu), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, u), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, ru), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, r), cellWay + orto * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, rd), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, d), cellWay + orto * grain, maxWay) Then
        Set OneStep = Nothing
            Exit Function
        End If
        If IsIntercourseWithFront(nozzle, newColl, NextCell(cell, ld), cellWay + diag * grain, maxWay) Then
            Set OneStep = Nothing
            Exit Function
        End If

    Next cell
    
    Set OneStep = newColl
    
    

End Function

Private Function IsIntercourseWithFront(ByRef nozzle As c_FireNozzleModel, ByRef coll As Collection, ByRef cell As c_Cell, _
                                        ByVal wayLen As Single, ByVal maxWay As Integer) As Boolean
'Проверяем путь пройденный до клетки и если он больше входящего, заменяем его
'Клетку пытаемся добавить в пришедшую коллекцию
    'Если клетка находится за пределами матрицы - выходим
    If Not IsInMatrix(cell) Then Exit Function
    
    'Проверяем, является ли ячейка открытым пространством
    If Not fireMatrixObject.GetOpenSpaceLayerValue(cell.x, cell.y) = CellState.csOpenSpace Then Exit Function
    
    'Если путь превышает предельное значение - не учитываем его
    If wayLen > maxWay Or calculatedCells(cell.x, cell.y) = 1 Then Exit Function
    
    'Если больше предлагаемого пути, то...
    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        'Проверяем, является ли клетка клеткой фронта пожара
    'ДОБАВИТЬ ПРОВЕРКУ МАКСИМАЛЬНОГО ГОРЕНИЯ!!!
        If fireMatrixObject.IsCellMaxFire(cell.x, cell.y) Then
            If Not fireMatrixObject.IsInner(cell.x, cell.y) Then
                'Добавляем для ствола значение точки контакта с фронтом пожара
'                Debug.Print "cell." & cell.x & ":" & cell.y & " команда на отрисовку!!!"
                Set nozzle.fireContactCell = cell
                
                'Выходим из функции возвратив Истина - зона работы ствола касается фронта пожара (найдена!)
                IsIntercourseWithFront = True
                Exit Function
            End If
        End If
        
        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
        matrix(cell.x, cell.y) = wayLen

        On Error Resume Next
        coll.Add cell, cell.x & ":" & cell.y
    End If
    
End Function

Private Function IsIntercourseWithFront2(ByRef nozzle As c_FireNozzleModel, ByRef cell As c_Cell) As Boolean
'Проверяем является ли клетка частью вронта пожара
    'Если клетка находится за пределами матрицы - выходим
    If Not IsInMatrix(cell) Then Exit Function
    
    'Проверяем, является ли ячейка открытым пространством
    If Not fireMatrixObject.GetOpenSpaceLayerValue(cell.x, cell.y) = CellState.csOpenSpace Then
        'Выходим из функции - найдена стена, но фронт не найден, ячейка точки касания фронта не задана
        IsIntercourseWithFront2 = True
        Exit Function
    End If
    
'    'Если больше предлагаемого пути, то...
'    If matrix(cell.x, cell.y) = 0 Or matrix(cell.x, cell.y) > wayLen Then
        'Проверяем, является ли клетка клеткой фронта пожара
    'ДОБАВИТЬ ПРОВЕРКУ МАКСИМАЛЬНОГО ГОРЕНИЯ!!!
        If fireMatrixObject.IsCellMaxFire(cell.x, cell.y) Then
            If Not fireMatrixObject.IsInner(cell.x, cell.y) Then
                'Добавляем для ствола значение точки контакта с фронтом пожара
'                Debug.Print "cell." & cell.x & ":" & cell.y & " команда на отрисовку!!!"
                Set nozzle.fireContactCell = cell
                
                'Выходим из функции возвратив Истина - зона работы ствола касается фронта пожара (найдена!)
                IsIntercourseWithFront2 = True
                Exit Function
            End If
        End If
        
'        'Ставим новое значение и пытаемся добавить в коллекция (если такая клетка в коллекции уже есть, ничего не происходит - через Error)
'        matrix(cell.x, cell.y) = wayLen

'        On Error Resume Next
'        coll.Add cell, cell.x & ":" & cell.y
'    End If
    
End Function




'-----------------------Инструменты-----------------------------------------
Private Function NextCell(ByRef cell As c_Cell, ByVal dir As Directions) As c_Cell
'Возвращает следующую клетку согласно указанного направления
    Select Case dir
        Case Is = Directions.l
            Set NextCell = NewCell(cell.x - 1, cell.y)
        Case Is = Directions.lu
            Set NextCell = NewCell(cell.x - 1, cell.y + 1)
        Case Is = Directions.u
            Set NextCell = NewCell(cell.x, cell.y + 1)
        Case Is = Directions.ru
            Set NextCell = NewCell(cell.x + 1, cell.y + 1)
        Case Is = Directions.r
            Set NextCell = NewCell(cell.x + 1, cell.y)
        Case Is = Directions.rd
            Set NextCell = NewCell(cell.x + 1, cell.y - 1)
        Case Is = Directions.d
            Set NextCell = NewCell(cell.x, cell.y - 1)
        Case Is = Directions.ld
            Set NextCell = NewCell(cell.x - 1, cell.y - 1)
    End Select
End Function

Private Function NewCell(ByVal x As Long, ByVal y As Long) As c_Cell
'Создаем новую клетку
Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = x
    cell.y = y
    
    Set NewCell = cell
End Function

Private Function IsInMatrix(ByRef cell As c_Cell) As Boolean
'Возвращает Истина, если клетка находится в пределах матрицы
    If cell.x <= 0 Or cell.y <= 0 Or cell.x > UBound(matrix, 1) Or cell.y > UBound(matrix, 2) Then
        IsInMatrix = False
    Else
        IsInMatrix = True
    End If
End Function


