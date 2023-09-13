VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExtSquareCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Const diag As Double = 0.1                                  ' Коэффициент мощности по диагонали
'Const orto As Double = 0.14142135623731                     ' Коэффициент мощности по вертикали и горизонтали
Const mmInInch As Double = 25.4

'Private matrix As c_MatrixExtSquare                         'Матрица окружающего пространства (упрощенная для расчета площади тушения)
Private matrix As c_Matrix                                  'Матрица окружающего пространства
Private extSquare As c_ExtinguishingSquare                  'Объект площади тушения

Public grain As Double              'размер зерна расчета (клетки)
'Private step As Integer             'Счетчик количества шагов (общий, без моделирвоания и с моделирвоанием)
'Private calcStep As Integer         'Счетчик количества раелизованных (ТОЛЬКО с моделированием) шагов
'Public distance As Single           'общая дистанция пройденная при моделирвоании (условная для идеального случая)
'Public time As Single               'общее время прошедшее при моделирвоании (условная для идеального случая)
'Public intenseNeed As Single        'Требуемая интенсивность подачи воды л/(с*м.кв.)

Private nozzleDemon As c_FireNozzlesDemon
    

'------------------------------------Класс демона рассчитывающего площадь тушения----------------------------------------------














Public Sub SetOpenSpaceLayer(ByRef modeller As c_Modeller)
'Устанавливаем ссылку на матрицу окружающего пространства
Dim openSpaceLayer As Variant
Dim shp As Visio.Shape
    
    openSpaceLayer = modeller.GetOpenSpaceLayer
    
    Set matrix = New c_Matrix
    matrix.CreateMatrix UBound(openSpaceLayer, 1), UBound(openSpaceLayer, 2)
    matrix.SetOpenSpace openSpaceLayer
    
Dim i As Long
Dim j As Long
    
    For i = 0 To UBound(openSpaceLayer, 1)
        For j = 0 To UBound(openSpaceLayer, 2)
            If openSpaceLayer(i, j) = 0 Then
'                Set shp = Application.ActivePage.DrawRectangle((i * grain) / mmInInch, (j * grain) / mmInInch, _
'                                                     (i * grain + grain) / mmInInch, (j * grain + grain) / mmInInch)
'                                                     shp.Cells("FillForegnd").Formula = 3
'                                                     shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("OpenSpace")
            End If
        Next j
    Next i
    
    grain = modeller.grain
End Sub




Public Sub RunDemon(ByRef shp As Visio.Shape)
'Запускаем демона
Dim fireShapeDemon As c_FireShapeDemon
Dim squareShapesCol As Collection
Dim extShp As Shape
Dim ExtSquareFormula As String

'    Set squareShapesCol = New Collection
    
    Set fireShapeDemon = New c_FireShapeDemon
    'Расчитываем сетку фигуры
    fireShapeDemon.setGrain(grain).SetFireShape(shp).RunDemon matrix
    'Расчитываем блоки зоны тушения
    FireSquareCalculate
    'Строим площадь тушения
    Set squareShapesCol = DrawExtSquareByDemon
    'Составляем стрку формулы для ссылки на размеры полученных площадей тушения
    For Each extShp In squareShapesCol
        If ExtSquareFormula = "" Then
            ExtSquareFormula = ExtSquareFormula & "IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        Else
            ExtSquareFormula = ExtSquareFormula & "+IF(ISERR(Sheet." & extShp.ID & "!Prop.ExtSquare),0,Sheet." & extShp.ID & "!Prop.ExtSquare)"
        End If
    Next extShp
    'Дополняем формулу округлением
    ExtSquareFormula = "Round(" & ExtSquareFormula & ",0)"
    'Присваиваем значению площади тушения рассматриваемой фигуры площади пожара значение сумм площадей тушения полченных фигур (формулу)
    shp.Cells("Prop.ExtFull").FormulaU = "Index(1, Prop.ExtFull.Format)"
    shp.Cells("Prop.ExtSquareT").FormulaU = ExtSquareFormula
End Sub





'----------------------------Расчет площадей тушения--------------------------
Public Sub FireSquareCalculate()
'Процедура получения блоков площади тушения
Dim nozzle As c_FireNozzleModel
Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
    
    Set extSquare = New c_ExtinguishingSquare
'    extSquare.intenseNeed = intenseNeed
    
    Set frontDemonImpl = New c_FrontDemon
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    frontDemonImpl.CreateMatrix matrix.DimensionX, matrix.DimensionY
    frontDemonImpl.grain = grain
    frontDemonImpl.SetStartState matrix.fireFrontCells
    
    extinguishingSquareDemon.CreateMatrix matrix
    extinguishingSquareDemon.setGrain = grain
    
    frontDemonImpl.RunDemon2 extSquare, matrix.fireFrontCells, matrix.fireFrontCells
'    extinguishingSquareDemon.RunDemon extSquare
    extinguishingSquareDemon.RunDemon extSquare
    
'    Dim cell As c_Cell
'    for each cell
    
    
    
    
    
'    Dim i As Integer
'    i = 0
'    For Each nozzle In nozzleCollection
'        If Not nozzle.fireContactCell Is Nothing Then
''            Debug.Assert i < 1
'
''            If i > 0 Then frontDemonImpl.DrawBlockExists extSquare, nozzle.fireContactCell
'            'Формируем коллекцию фрагментов площади фронта пожара (и блоков тушения на их основе)
'            frontDemonImpl.RunDemon extSquare, matrix.fireFrontCells, nozzle
'
'            'Рассчитываем размер каждого из блоков
'            extinguishingSquareDemon.RunDemon extSquare
'
'            i = i + 1
'        End If
'    Next nozzle
    
    
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

Public Function DrawExtSquareByDemon() As Collection
'Команда демону отрисовать площадь тушения
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon
Dim fireConturs As Visio.Selection
Dim shp As Visio.Shape
'Dim squareShapesCol As Collection
'
'    Set squareShapesCol = New Collection
    
    ClearLayer "ExtSquare"
    
    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon

    extinguishingSquareDemon.setGrain = grain
    Set DrawExtSquareByDemon = extinguishingSquareDemon.DrawFireSquare(extSquare)
    
    
    
    'Объединяем полученные фигуры площади тушения в окончательную на текущем шаге
'    On Error Resume Next
'    Set fireConturs = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "ExtSquare")
'    If fireConturs.Count > 0 Then
'        fireConturs.Combine
'        Set shp = Application.ActiveWindow.Selection(1)
'        shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("ExtSquare")
'    End If
End Function

