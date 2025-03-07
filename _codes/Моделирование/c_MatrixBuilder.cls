VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MatrixBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------Класс для получения матрицы----------------------------------------------------------------------------------
Option Explicit


Private GRAIN As Integer        'Размер зерна в мм
Const mmInInch As Double = 25.4
Private grain_i As Double
Private currentCells As Collection

Private form As Object

Private bordreRect As Visio.Shape

Public maskBuilder As c_MaskBuilder

Public matrix As Variant




Public Sub SetForm(ByRef frm As Object)
'Указываем ссылку на внешнюю форму в которую будет выводиться отчет о процессе запекания формы
    Set form = frm
End Sub


Public Function NewMatrix(ByVal grainVal As Integer) As Variant()
'Функция создает и возвращает новую матрицу объекта для расчета
'grain - размер зерна в мм
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer
Dim areaBuilder As C_MatrixAreaBuilder
Dim shp As Visio.Shape
Dim perc As Double                      'Процент открытия двери
Dim PinX As Long
Dim PinY As Long

    GRAIN = grainVal
    grain_i = GRAIN / mmInInch

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / GRAIN
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / GRAIN
    Debug.Print "Размер матрицы " & xCount * yCount
    
    ReDim matrix(xCount, yCount)
    
    'Проверяем наличие контурной фигуры
    TryGetShape bordreRect, "User.IndexPers:1001"
    
    
    'Включаем цвет подписи в форме отчета хода расчета
    form.lblMatrixIsBaked.ForeColor = vbBlack
    
    '---Заполняем матрицу нолями
    FillMatrixZero

    '---Заполняем матрицу неучетным пространством в случае, если указана расчетная зона
    If Not bordreRect Is Nothing Then
        Set areaBuilder = New C_MatrixAreaBuilder
            form.lblMatrixIsBaked.Caption = "Оценка расчетной зоны, подождите"
            form.lblMatrixIsBaked.ForeColor = vbBlack
            DoEvents

            FillMatrixWalls
            matrix = areaBuilder.SetForm(form).FillArea(matrix, bordreRect, GRAIN)
'            matrix = areaBuilder.FillArea(matrix, bordreRect, GRAIN)
        Set areaBuilder = Nothing
    End If

    '---Рекурсивная идентификация контура стен
    Debug.Print "Определение контура стен"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
'                Application.ActiveWindow.Select shp, visSelect
                Set currentCells = New Collection
                PinX = CellVal(shp, "PinX", visMillimeters)
                PinY = CellVal(shp, "PinY", visMillimeters)
'                Patch PinX, PinY
                Attack shp, PinX, PinY
                
                form.AddCheckedSize 1, 2
            End If
        Next shp
    Else
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
                If bordreRect.SpatialRelation(shp, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                    Set currentCells = New Collection
'                    Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters)
                    PinX = CellVal(shp, "PinX", visMillimeters)
                    PinY = CellVal(shp, "PinY", visMillimeters)
'                    Patch PinX, PinY
                    Attack shp, PinX, PinY
                    
                    form.AddCheckedSize 1, 2
                End If
            End If
        Next shp
    End If

    
    '---Рекурсивная идентификация контура дверей
    Debug.Print "Определение контура окон"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isDoor(shp) Then
'                Application.ActiveWindow.Select shp, visSelect
                perc = CloseDoor(shp)
                
                Set currentCells = New Collection
                Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters), csOpenSpace
                
                OpenDoor shp, perc
                
                form.AddCheckedSize 1, 3
            End If
        Next shp
    Else
        For Each shp In Application.ActivePage.Shapes
            If PFB_isDoor(shp) Then
                If bordreRect.SpatialRelation(shp, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                    perc = CloseDoor(shp)
                    
                    Set currentCells = New Collection
                    Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters), csOpenSpace
                    
                    OpenDoor shp, perc
                    
                    form.AddCheckedSize 1, 3
                End If
            End If
        Next shp
    End If


    '---Возвращаем результат - матрицу заполненную данными о ограждающих конструкциях
    NewMatrix = matrix

End Function

Public Sub Attack(ByRef shp As Visio.Shape, ByVal X As Double, ByVal Y As Double, Optional ByVal state As CellState = CellState.csWall)
    
Dim radiuss As Double
Dim x_i As Integer
Dim y_i As Integer
Dim pnt As c_Point

    On Error GoTo ErrRange
    
    ' Приведение к размерности матрицы
    X = Int(X / GRAIN) * GRAIN + GRAIN / 2
    Y = Int(Y / GRAIN) * GRAIN + GRAIN / 2
    
    x_i = X / mmInInch
    y_i = Y / mmInInch
'    grain_i = grain / mmInInch
    
    If X < 0 Or Y < 0 Or X > UBound(matrix, 1) * GRAIN Or Y > UBound(matrix, 2) * GRAIN Then
        Exit Sub
    End If
    
'    DoEvents
    
    radiuss = grain_i * 0.75
'    radiuss = grain_i

    Set pnt = New c_Point
    pnt.SetData Int(X / GRAIN), Int(Y / GRAIN)
    If IsInCollection(currentCells, pnt) Then
        Exit Sub
    End If

'                'Отрисовываем квадрат
'                Dim shpT As Visio.Shape
'                Set shpT = Application.ActivePage.DrawRectangle(x / mmInInch, y / mmInInch, _
'                        (x + grain) / mmInInch, (y + grain) / mmInInch)
'                shpT.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("test")
'                shpT.Cells("LinePattern").FormulaU = "0"
'                shpT.Cells("FillForegnd").FormulaU = "2"
    If shp.HitTest(x_i, y_i, radiuss) Then
'        Debug.Print "c_MatrixBuilder->Attack->shp.HitTest: Уточнить радус и порадок расета - не срабатывает для тонких стен"
        
        currentCells.Add pnt
        matrix(Int(X / GRAIN), Int(Y / GRAIN)) = state
        
        Attack shp, X - GRAIN, Y, state
        Attack shp, X - GRAIN, Y + GRAIN, state
        Attack shp, GRAIN, Y + GRAIN, state
        Attack shp, X + GRAIN, Y + GRAIN, state
        Attack shp, X + GRAIN, Y, state
        Attack shp, X + GRAIN, Y - GRAIN, state
        Attack shp, GRAIN, Y - GRAIN, state
        Attack shp, X - GRAIN, Y - GRAIN, state
    End If
    
Exit Sub
ErrRange:
    If Err.number = 9 Then
        Exit Sub
    Else
        Debug.Print Err.description
    End If
End Sub

Private Sub CheckEnvironment(Optional ByRef rect As c_Rect)
'Процедура определяет конфигурацию пространства
Dim startRectSizeHalf As Long
    
    'Рекурсивно анализируем схему
    '---определяем размер стартового квадрата для поиска стен
        startRectSizeHalf = GetMaximumRectSize / 2   'Половина потому что в самом начале и радиус и количество зерен равны
        'Задаем для внешней формы значение размера матрицы
        form.SetMatrixSize (startRectSizeHalf * 2) ^ 2
    '---Запускаем рекурсивный анализ
        CheckRect startRectSizeHalf, startRectSizeHalf, startRectSizeHalf
End Sub


Private Function IsOpenSpace(X As Long, Y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской стен, и наоборот, если соприкасается
    IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(X * GRAIN / mmInInch, Y * GRAIN / mmInInch, tolerance / mmInInch) = 0
End Function

Private Function IsTotalWall(X As Long, Y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской свободного пространства, и наоборот, если соприкасается
    IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(X * GRAIN / mmInInch, Y * GRAIN / mmInInch, tolerance / mmInInch) = 0
End Function

Private Sub Patch(X, Y, Optional ByVal state As CellState = CellState.csOpenSpace)
' Процедура добавляет в точку матрицы с координатами x, y накладку из клеток в которых установлено значение CellState
Dim xx As Integer
Dim yy As Integer

    xx = Int(X / GRAIN)
    yy = Int(Y / GRAIN)

'                'Отрисовываем квадрат
'                Dim shpT As Visio.Shape
'                Set shpT = Application.ActivePage.DrawRectangle(x / mmInInch, y / mmInInch, _
'                        (x + grain) / mmInInch, (y + grain) / mmInInch)
'                shpT.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("test")
'                shpT.Cells("LinePattern").FormulaU = "0"
'                shpT.Cells("FillForegnd").FormulaU = "2"

    matrix(xx - 1, yy - 1) = state
    matrix(xx - 1, yy) = state
    matrix(xx - 1, yy + 1) = state
    matrix(xx, yy - 1) = state
    matrix(xx, yy) = state
    matrix(xx, yy + 1) = state
    matrix(xx + 1, yy - 1) = state
    matrix(xx + 1, yy) = state
    matrix(xx + 1, yy + 1) = state
End Sub




'------------------Рекурсивное составление матрицы------------------------------------------------------
Private Sub CheckRect(ByVal X As Long, ByVal Y As Long, ByVal radiuss As Integer)
'Здесь radiuss - в зернах!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       'Модификатор радиуса - нужен, что бы гарантировать проверку всей площади квадрата - переводит длину пришедшего радиуса (являющегося половиной длины квадрата) к радиусу круга описывающего квадрат, в идеале равен 1,4142135623731

    On Error GoTo EX
               
               
'1 Проверяем утверждение, что в радиусе нет стен
    '---Если стен нет, оставляем для всей клетки 0
        If IsOpenSpace(X, Y, radiuss * radiussModificator * GRAIN) Then
            form.AddCheckedSize (radiuss * 2) ^ 2

'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                    (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'            shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
            DoEvents

            Exit Sub
        End If
        
'2 Если стены есть, проверяем меньшие квадраты
'3 проверяем утверждение, что в радиусе нет открытого пространства
    '---Если стены есть, проверяем более детально
        If IsTotalWall(X, Y, radiuss * radiussModificator * GRAIN) Then
        'Если при этом в радиусе нет открытого пространства, ставим для всех входящих точек значение стены (1)
            SetToAllCells X, Y, radiuss, csWall
            form.AddCheckedSize (radiuss * 2) ^ 2
            
'                Dim shp As Visio.Shape
'                Set shp = Application.ActivePage.DrawRectangle((X - radiuss) * grain / mmInInch, (Y - radiuss) * grain / mmInInch, _
'                        (X + radiuss) * grain / mmInInch, (Y + radiuss) * grain / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
                DoEvents
        Else
            If radiuss = 1 Then
                form.AddCheckedSize 1
                
                'Проверяем все входящие клетки, т.к. при радиусе = 1 в клетку входит четыре атомарных
'                radiussHalf = 0.5
'                '1 - вверх влево
'                    checkSingleCell x - 0.5, y + 0.5
'                '2 - вверх вправо
'                    checkSingleCell x + 0.5, y + 0.5
'                '3 - вниз вправо
'                    checkSingleCell x + 0.5, y - 0.5
'                '4 - вниз влево
'                    checkSingleCell x - 0.5, y - 0.5
                '1 - вверх влево
                    CheckSingleCell X, Y, False, True
                '2 - вверх вправо
                    CheckSingleCell X, Y, True, True
                '3 - вниз вправо
                    CheckSingleCell X, Y, True, False
                '4 - вниз влево
                    CheckSingleCell X, Y, False, False
'                Проверяем текущую клетку, т.к. при радиусе = 0 она является атомарной и далее не делится
'                checkSingleCell x + 0.5, y + 0.5
                Exit Sub
'            End If
            Else
            
            'Иначе рекурсивно вызываем проверку четырех входящих квадратов
            radiussHalf = radiuss * 0.5
            '1 - вверх влево
                CheckRect X - radiussHalf, Y + radiussHalf, radiussHalf
            '2 - вверх вправо
                CheckRect X + radiussHalf, Y + radiussHalf, radiussHalf
            '3 - вниз вправо
                CheckRect X + radiussHalf, Y - radiussHalf, radiussHalf
            '4 - вниз влево
                CheckRect X - radiussHalf, Y - radiussHalf, radiussHalf
            
            End If
        End If
Exit Sub
EX:
    Debug.Print "x=" & X & ", y=" & Y & ", radiuss=" & radiuss
End Sub



Private Sub CheckSingleCell(ByVal X As Integer, ByVal Y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean, Optional tolerance As Double = 0)
'Проверяем наличие стен для единичной (атомарной ячейки)
Dim xAct As Double
Dim yAct As Double
Const radiussModificator As Double = 1.4142

    xAct = X + 0.5 + Not dirRight
    yAct = Y + 0.5 + Not dirUp
    
    If Not maskBuilder.pO_MaskShape.HitTest(xAct * GRAIN / mmInInch, yAct * GRAIN / mmInInch, radiussModificator * GRAIN * 0.5 / mmInInch) = 0 Then
        If dirRight = True And dirUp = True Then matrix(X, Y) = CellState.csWall
        If dirRight = False And dirUp = True Then matrix(X - 1, Y) = CellState.csWall
        If dirRight = True And dirUp = False Then matrix(X, Y - 1) = CellState.csWall
        If dirRight = False And dirUp = False Then matrix(X - 1, Y - 1) = CellState.csWall
        
        
'        Debug.Print x * grain & ":" & y * grain & " Стена!"
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    Else
'                    X = X + Not dirRight
'                    Y = Y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(X * grain / mmInInch, Y * grain / mmInInch, _
'                        (X + 1) * grain / mmInInch, (Y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
    End If

End Sub

Private Sub SetToAllCells(ByVal X As Long, ByVal Y As Long, ByVal size As Long, ByVal value As CellState)
'Устанавливаем для всех клеток указанного диапазона size с центром в точке x,y значение value
Dim i As Long
Dim j As Long
    
    For i = X - size To X + size
        For j = Y - size To Y + size
            matrix(i, j) = value
        Next j
    Next i
End Sub

Private Function GetMaximumRectSize() As Long
'Возвращаем максимально возможный квадрат полностью покрывающий схему и при этом получающий без остатка любой результат умножения на 2
Dim grains As Long
Dim pageWidth As Double
Dim pageHeight As Double

    grains = 1
    
    pageWidth = Application.ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters)
    pageHeight = Application.ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters)
    
    Do While grains * GRAIN < pageWidth Or grains * GRAIN < pageHeight
        grains = grains * 2
    Loop
    
GetMaximumRectSize = grains
End Function

'------------------Рекурсивное составление матрицы------------------------------------------------------

Private Sub FillMatrixZero()
'Заполняем матрицу нолями
Dim X As Long
Dim Y As Long
    
    For Y = 0 To UBound(matrix, 1)
        For X = 0 To UBound(matrix, 2)
            matrix(Y, X) = CellState.csOpenSpace
        Next X
    Next Y
End Sub

Private Sub FillMatrixWalls()
'Заполняем матрицу нолями
Dim X As Long
Dim Y As Long
    
    For Y = 0 To UBound(matrix, 1)
        For X = 0 To UBound(matrix, 2)
            matrix(Y, X) = CellState.csWall
        Next X
    Next Y
End Sub

'------------------Функции-----------------------
Private Function CloseDoor(ByRef shp As Visio.Shape) As Double
Dim perc As Double
    
    CloseDoor = CellVal(shp, "Prop.VisDoorOpenPercent")
    SetCellVal shp, "Prop.VisDoorOpenPercent", 0
    
End Function

Private Sub OpenDoor(ByRef shp As Visio.Shape, Optional ByVal perc As Double = 50)
     SetCellVal shp, "Prop.VisDoorOpenPercent", perc
End Sub
