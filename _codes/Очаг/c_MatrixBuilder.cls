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

Private form As Object

Private bordreRect As Visio.Shape

Public maskBuilder As c_MaskBuilder

Public matrix As Variant



'Public Function GetMatrix(ByRef refMatrix() As Variant) As Boolean
''Если матрица уже была создана - возвращает True и присваивает ее значение ссылке, если нет, то возвращает false
'    If matrix = Empty Then
'        GetMatrix = False
'    Else
'        refMatrix = matrix
'        GetMatrix = True
'    End If
'End Function

Public Sub SetForm(ByRef frm As Object)
'Указываем ссылку на внешнюю форму в которую будет выводиться отчет о процессе запекания формы
    Set form = frm
End Sub


Public Function NewMatrix(ByVal grainVal As Integer, Optional ByRef rect As c_Rect) As Variant()
'Функция создает и возвращает новую матрицу объекта для расчета
'grain - размер зерна в мм
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer

    grain = grainVal

    xCount = ActivePage.PageSheet.Cells("PageWidth").Result(visMillimeters) / grain
    yCount = ActivePage.PageSheet.Cells("PageHeight").Result(visMillimeters) / grain
    Debug.Print "Размер матрицы " & xCount * yCount
    
    ReDim matrix(xCount, yCount)
       
    'Проверяем наличие контурной фигуры
    TryGetShape bordreRect, "User.IndexPers:1001"
       
    'Определяем конфигурацию матрицы (стены, зоны горения и т.д.)
    Set maskBuilder = New c_MaskBuilder
    '---Создаем маску
    maskBuilder.s_makeMask
        '---Если маска стен не была создана в результате отсутсвия таковых - возвращаем чисутю матрицу
        If maskBuilder.pO_MaskShape Is Nothing Then
            '---Возвращаем результат - матрицу заполненную данными о ограждающих конструкциях
            NewMatrix = matrix
            
            '---Уничтожаем построитель масок
            Set maskBuilder = Nothing
            Exit Function
        End If
    '---Создаем маску свободного пространства
    maskBuilder.s_makeOpenSpaceMask

    '---Определяем окружение
    CheckEnvironment


    '---Возвращаем результат - матрицу заполненную данными о ограждающих конструкциях
    NewMatrix = matrix

    '---Удаляем маски
    maskBuilder.pO_MaskShape.Delete
    maskBuilder.pO_FreespaceMaskShape.Delete
    
    '---Уничтожаем построитель масок
    Set maskBuilder = Nothing

End Function

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
    'Доделать учет ограничивающей фигуры!!!
    If bordreRect Is Nothing Then
        'Снаружи только маски стен
        IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
    Else
        'Снаружи и маски стен и ограничивающей фигуры
'        flag1 = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
'        flag2 = bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
'        IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0 And _
'                      bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) > 0
        
        'Если за пределами, то не учитывается!
'        IsOpenSpace = bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) > 0
        If bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0 Then
            IsOpenSpace = False
'            IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
        Else
            IsOpenSpace = True
            IsOpenSpace = maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
        End If
'        IsOpenSpace = flag1 Or flag2
    End If
    
End Function

Private Function IsTotalWall(x As Long, y As Long, tolerance As Double) As Boolean
'Функция возвращает Истина, если указанная точка не соприкасается с маской свободного пространства, и наоборот, если соприкасается
'    IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
    
    If bordreRect Is Nothing Then
        'Снаружи только маски стен
        IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
    Else
        'Снаружи и маски стен и ограничивающей фигуры
        'Если за пределами, то не учитывается!
'        IsTotalWall = bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
        If bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0 Then
            IsTotalWall = True
'            IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
        Else
            IsTotalWall = False
            IsTotalWall = maskBuilder.pO_FreespaceMaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0
        End If

    End If
End Function




'------------------Рекурсивное составление матрицы------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'Здесь radiuss - в зернах!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       'Модификатор радиуса - нужен, что бы гарантировать проверку всей площади квадрата - переводит длину пришедшего радиуса (являющегося половиной длины квадрата) к радиусу круга описывающего квадрат, в идеале равен 1,4142135623731

    On Error GoTo EX
               
''0 Проверяем утверждение, что в радиусе нет нерассчитываемых зон
'    If Not bordreRect Is Nothing Then
'        If bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) > 1 Then
'            form.AddCheckedSize (radiuss * 2) ^ 2
'
''            Exit Sub
'        End If
'
'    End If
               
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
                    checkSingleCell x, y, False, True
                '2 - вверх вправо
                    checkSingleCell x, y, True, True
                '3 - вниз вправо
                    checkSingleCell x, y, True, False
                '4 - вниз влево
                    checkSingleCell x, y, False, False
'                Проверяем текущую клетку, т.к. при радиусе = 0 она является атомарной и далее не делится
'                checkSingleCell x + 0.5, y + 0.5
                Exit Sub
            End If
            
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
Exit Sub
EX:
    Debug.Print "x=" & x & ", y=" & y & ", radiuss=" & radiuss
End Sub



Private Sub checkSingleCell(ByVal x As Integer, ByVal y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean)
'Проверяем наличие стен для единичной (атомарной ячейки)
Dim xAct As Double
Dim yAct As Double
Const radiussModificator As Double = 1.4142

    xAct = x + 0.5 + Not dirRight
    yAct = y + 0.5 + Not dirUp

'    If Not maskBuilder.pO_MaskShape.HitTest(x * grain / mmInInch, y * grain / mmInInch, radiussModificator * grain * 0.5 / mmInInch) = 0 Then
    If Not bordreRect Is Nothing Then
        If bordreRect.HitTest(x * grain / mmInInch, y * grain / mmInInch, tolerance / mmInInch) = 0 Then
            If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csWall
            If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csWall
            If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csWall
            If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csWall
        End If
        Exit Sub
    End If
    
    
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




