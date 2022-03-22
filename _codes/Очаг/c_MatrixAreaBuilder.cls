VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_MatrixAreaBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------Класс для получения матрицы свободных пространств для расчета----------------------------------------------------------------------------------
Option Explicit


Private grain As Integer        'Размер зерна в мм
Const mmInInch As Double = 25.4

'Private form As Object

Private bordreRect As Visio.Shape
Private scanner As Visio.Shape      'Фигура-сканер которая будет оценивать пространство

Public matrix As Variant




'Public Sub SetForm(ByRef frm As Object)
''Указываем ссылку на внешнюю форму в которую будет выводиться отчет о процессе запекания формы
'    Set form = frm
'End Sub


Public Function FillArea(ByRef matrix_a As Variant, ByRef bordreRect_a As Visio.Shape, ByVal grain_a As Integer) As Variant
'Функция создает и возвращает матрицу расчетного пространства (того, которое будет участвовать в расчете)
'grain - размер зерна в мм
Dim xCount As Long
Dim yCount As Long
Dim arr() As Integer

    grain = grain_a
    matrix = matrix_a
    Set bordreRect = bordreRect_a
    
    '---Создаме фигуру-сканер
    NewScaner
    
    '---Определяем окружение
    CheckEnvironment


    '---Возвращаем результат - матрицу заполненную данными о расчетном пространстве
    FillArea = matrix
    
    'Удаляем фигуру сканер
    scanner.Delete

End Function

Private Sub CheckEnvironment()
'Процедура определяет конфигурацию пространства
Dim startRectSizeHalf As Long
    
    'Рекурсивно анализируем схему
    '---определяем размер стартового квадрата для поиска стен
        startRectSizeHalf = GetMaximumRectSize / 2   'Половина потому что в самом начале и радиус и количество зерен равны
'        'Задаем для внешней формы значение размера матрицы
'        form.SetMatrixSize (startRectSizeHalf * 2) ^ 2
    '---Запускаем рекурсивный анализ
        CheckRect startRectSizeHalf, startRectSizeHalf, startRectSizeHalf
End Sub





'------------------Рекурсивное составление матрицы------------------------------------------------------
Private Sub CheckRect(ByVal x As Long, ByVal y As Long, ByVal radiuss As Integer)
'Здесь radiuss - в зернах!!!!
Dim radiussHalf As Integer
Const radiussModificator As Double = 1.4142       'Модификатор радиуса - нужен, что бы гарантировать проверку всей площади квадрата - переводит длину пришедшего радиуса (являющегося половиной длины квадрата) к радиусу круга описывающего квадрат, в идеале равен 1,4142135623731
Dim scanResult As Byte

    On Error GoTo EX
    
    
    SanerPos x, y, radiuss * 2
    scanResult = Scan
'1 Проверяем утверждение, что в радиусе нет расчетной области (везде - РО)
    '---Если стен нет, оставляем для всей клетки 0
        If scanResult = 0 Then
'            form.AddCheckedSize (radiuss * 2) ^ 2
'            form.Refresh
'            Dim shp As Visio.Shape
'            Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                    (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'            shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'            shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'
'            DoEvents

            Exit Sub
        End If
        
'2 Если стены есть, проверяем меньшие квадраты
'3 проверяем утверждение, что в радиусе нет открытого пространства
    '---Если стены есть, проверяем более детально
        If scanResult = 2 Then
        'Если при этом в радиусе нет открытого пространства, ставим для всех входящих точек значение стены (1)
            SetToAllCells x, y, radiuss, CellState.csOpenSpace
'            form.AddCheckedSize (radiuss * 2) ^ 2
            
'                Dim shp As Visio.Shape
'                Set shp = Application.ActivePage.DrawRectangle((x - radiuss) * grain / mmInInch, (y - radiuss) * grain / mmInInch, _
'                        (x + radiuss) * grain / mmInInch, (y + radiuss) * grain / mmInInch)
'                shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'                shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
''
'                DoEvents
        Else
            If radiuss = 1 Then
'                form.AddCheckedSize 1
                
                'Проверяем все входящие клетки, т.к. при радиусе = 1 в клетку входит четыре атомарных
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



Private Sub checkSingleCell(ByVal x As Integer, ByVal y As Integer, ByVal dirRight As Boolean, ByVal dirUp As Boolean, Optional tolerance As Double = 0)
'Проверяем наличие стен для единичной (атомарной ячейки)
Dim scanResult As Byte

    SanerPos x, y
    scanResult = Scan
    
    If scanResult = 2 Then
        If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csOpenSpace
        If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csOpenSpace
        If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csOpenSpace
        If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csOpenSpace
        
        
'        Debug.Print x & ":" & y & " Открытое пространство!"
'        Debug.Print x * grain & ":" & y * grain & " Стена!"
'                    x = x + Not dirRight
'                    y = y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                        (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
'                    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("Пространство")
'                    shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.1 pt"
'    ElseIf scanResult = 2 Then
'        If dirRight = True And dirUp = True Then matrix(x, y) = CellState.csOpenSpace
'        If dirRight = False And dirUp = True Then matrix(x - 1, y) = CellState.csOpenSpace
'        If dirRight = True And dirUp = False Then matrix(x, y - 1) = CellState.csOpenSpace
'        If dirRight = False And dirUp = False Then matrix(x - 1, y - 1) = CellState.csOpenSpace
    
    
'                    x = x + Not dirRight
'                    y = y + Not dirUp
'                    Dim shp As Visio.Shape
'                    Set shp = Application.ActivePage.DrawRectangle(x * grain / mmInInch, y * grain / mmInInch, _
'                        (x + 1) * grain / mmInInch, (y + 1) * grain / mmInInch)
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



'------------------Функции сканера--------------------------
Public Sub NewScaner()
    Set scanner = Application.ActivePage.DrawRectangle(0, 0, 0, 0)
End Sub

Public Sub SanerPos(ByVal x As Double, ByVal y As Double, Optional ByVal size As Double = 1)
    SetCellVal scanner, "PinX", x * grain / mmInInch
    SetCellVal scanner, "PinY", y * grain / mmInInch
    SetCellVal scanner, "Width", size * grain / mmInInch
    SetCellVal scanner, "Height", size * grain / mmInInch
End Sub

Public Function Scan() As Byte
'0 - scanner не входит в bordreRect, 1 - накладывается, 2 - полностью укладывается
    Scan = bordreRect.SpatialRelation(scanner, 0, VisSpatialRelationFlags.visSpatialFrontToBack + VisSpatialRelationFlags.visSpatialIncludeHidden)
End Function
