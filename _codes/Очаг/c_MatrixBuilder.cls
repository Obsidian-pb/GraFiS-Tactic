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


Private grain As Integer        'Размер зерна в мм
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

    grain = grainVal
    grain_i = grain / mmInInch

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
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
'            form.lblMatrixIsBaked.ForeColor = vbBlack
            DoEvents

            FillMatrixWalls
            matrix = areaBuilder.SetForm(form).FillArea(matrix, bordreRect, grain)
        Set areaBuilder = Nothing
    End If

    '---Рекурсивная идентификация контура стен
    form.lblMatrixIsBaked.Caption = "Определение контура стен"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
                Set currentCells = New Collection
                Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters)
                
                form.AddCheckedSize 1, 2
            End If
        Next shp
    Else
        For Each shp In Application.ActivePage.Shapes
            If PFB_isWall(shp) Then
                If bordreRect.SpatialRelation(shp, 0, VisSpatialRelationFlags.visSpatialFrontToBack) > 0 Then
                    Set currentCells = New Collection
                    Attack shp, CellVal(shp, "PinX", visMillimeters), CellVal(shp, "PinY", visMillimeters)
                    
                    form.AddCheckedSize 1, 2
                End If
            End If
        Next shp
    End If

    
    '---Рекурсивная идентификация контура дверей
    form.lblMatrixIsBaked.Caption = "Определение контура окон"
    form.SetMatrixSize Application.ActivePage.Shapes.Count
'    DoEvents
    If bordreRect Is Nothing Then
        For Each shp In Application.ActivePage.Shapes
            If PFB_isDoor(shp) Then
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

Public Sub Attack(ByRef shp As Visio.Shape, ByVal x As Double, ByVal y As Double, Optional ByVal state As CellState = CellState.csWall)
    
Dim radiuss As Double
Dim x_i As Integer
Dim y_i As Integer
Dim pnt As c_Point

    On Error GoTo ErrRange
    
'    x = Int(x / grain) mmInInch
'    y = Int(x / grain)
    x_i = x / mmInInch
    y_i = y / mmInInch
'    grain_i = grain / mmInInch
    
    If x < 0 Or y < 0 Or x > UBound(matrix, 1) * grain Or y > UBound(matrix, 2) * grain Then
        Exit Sub
    End If
    
'    DoEvents
    
    radiuss = grain_i * 0.5

    Set pnt = New c_Point
    pnt.SetData Int(x / grain), Int(y / grain)
    If IsInCollection(currentCells, pnt) Then Exit Sub

    
    If shp.HitTest(x_i, y_i, radiuss) Then
        Debug.Print "c_MatrixBuilder->Attack->shp.HitTest: Уточнить радус и порадок расета - не срабатывает для тонких стен"
        
        currentCells.Add pnt
        matrix(Int(x / grain), Int(y / grain)) = state
        
        Attack shp, x - grain, y, state
        Attack shp, x - grain, y + grain, state
        Attack shp, grain, y + grain, state
        Attack shp, x + grain, y + grain, state
        Attack shp, x + grain, y, state
        Attack shp, x + grain, y - grain, state
        Attack shp, grain, y - grain, state
        Attack shp, x - grain, y - grain, state
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


Private Function IsOpenSpace(x As Long, y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской стен, и наоборот, если соприкасается
    IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function

Private Function IsTotalWall(x As Long, y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской свободного пространства, и наоборот, если соприкасается
    IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
End Function




'------------------Рекурсивное составление матрицы------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'Здесь radiuss - в зернах!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       'Модификатор радиуса - нужен, что бы гарантировать проверку всей площади квадрата - переводит длину пришедшего радиуса (являющегося половиной длины квадрата) к радиусу круга описывающего квадрат, в идеале равен 1,4142135623731

    On Error GoTo EX
               
               
'1 Проверяем утверждение, что в радиусе нет стен
    '---Если стен нет, оставляем для всей клетки 0
        If IsOpenSpace(x, y, radiuss * radiussModificator * grain) Then
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
        If IsTotalWall(x, y, radiuss * radiussModificator * grain) Then
        'Если при этом в радиусе нет открытого пространства, ставим для всех входящих точек значение стены (1)
            SetToAllCells x, y, radiuss, csWall
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
                    CheckSingleCell x, y, False, True
                '2 - вверх вправо
                    CheckSingleCell x, y, True, True
                '3 - вниз вправо
                    CheckSingleCell x, y, True, False
                '4 - вниз влево
                    CheckSingleCell x, y, False, False
'                Проверяем текущую клетку, т.к. при радиусе = 0 она является атомарной и далее не делится
'                checkSingleCell x + 0.5, y + 0.5
                Exit Sub
'            End If
            Else
            
            'Иначе рекурсивно вызываем проверку четырех входящих квадратов
            radiussHalf = radiuss * 0.5
            '1 - вверх влево
                CheckRect x - radiussHalf, y + radiussHalf, radiussHalf
            '2 - вверх вправо
                CheckRect x + radiussHalf, y + radiussHalf, radiussHalf
            '3 - вниз вправо
                CheckRect x + radiussHalf, y - radiussHalf, radiussHalf
            '4 - вниз влево
                CheckRect x - radiussHalf, y - radiussHalf, radiussHalf
            
            End If
        End If
Exit Sub
EX:
    Debug.Print "x=" & x & ", y=" & y & ", radiuss=" & radiuss
End Sub



Private Sub CheckSingleCell(ByVal x As Integer, ByVal y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean, Optional tolerance As Double = 0)
'Проверяем наличие стен для единичной (атомарной ячейки)
Dim xAct As Double
Dim yAct As Double
Const radiussModificator As Double = 1.4142

    xAct = x + 0.5 + Not dirRight
    yAct = y + 0.5 + Not dirUp
    
    If Not maskBuilder.pO_MaskShape.HitTest(xAct * grain / mmInInch, yAct * grain / mmInInch, radiussModificator * grain * 0.5 / mmInInch) = 0 Then
        If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csWall
        If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csWall
        If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csWall
        If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csWall
        
        
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

Private Sub SetToAllCells(ByVal x As Long, ByVal y As Long, ByVal size As Long, ByVal value As CellState)
'Устанавливаем для всех клеток указанного диапазона size с центром в точке x,y значение value
Dim i As Long
Dim j As Long
    
    For i = x - size To x + size
        For j = y - size To y + size
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
    
    Do While grains * grain < pageWidth Or grains * grain < pageHeight
        grains = grains * 2
    Loop
    
GetMaximumRectSize = grains
End Function

'------------------Рекурсивное составление матрицы------------------------------------------------------

Private Sub FillMatrixZero()
'Заполняем матрицу нолями
Dim x As Long
Dim y As Long
    
    For y = 0 To UBound(matrix, 1)
        For x = 0 To UBound(matrix, 2)
            matrix(y, x) = CellState.csOpenSpace
        Next x
    Next y
End Sub

Private Sub FillMatrixWalls()
'Заполняем матрицу нолями
Dim x As Long
Dim y As Long
    
    For y = 0 To UBound(matrix, 1)
        For x = 0 To UBound(matrix, 2)
            matrix(y, x) = CellState.csWall
        Next x
    Next y
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
