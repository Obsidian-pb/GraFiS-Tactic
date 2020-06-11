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
    
End Sub




Public Sub RunDemon()
'Запускаем демона
Dim fireShapeDemon As c_FireShapeDemon
Dim shp As Visio.Shape
'    Set matrix = New c_MatrixExtSquare
'    Set shp = Application.ActivePage.Shapes.ItemFromID(35)

'    Set shp = Application.ActivePage.Shapes.ItemFromID(326)
'    Set shp = Application.ActivePage.Shapes.ItemFromID(189)
'    Set shp = Application.ActivePage.Shapes.ItemFromID(187)
'    Set shp = Application.ActivePage.Shapes.ItemFromID(351)
'    Set shp = Application.ActivePage.Shapes.ItemFromID(94)
    Set shp = Application.ActivePage.Shapes.ItemFromID(15)
    
'    On Error GoTo EX
'    Set shp = Application.ActiveWindow.Selection(1)
    
    Set fireShapeDemon = New c_FireShapeDemon
    fireShapeDemon.setGrain(grain).SetFireShape(shp).RunDemon matrix
    
    FireSquareCalculate
    
    DrawExtSquareByDemon
    
    
    

EX:
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

Public Sub DrawExtSquareByDemon()
'Команда демону отрисовать площадь тушения
'Dim frontDemonImpl As c_FrontDemon
Dim extinguishingSquareDemon As c_ExtinguishSquareDemon

    Set extinguishingSquareDemon = New c_ExtinguishSquareDemon
    
    extinguishingSquareDemon.setGrain = grain
    extinguishingSquareDemon.DrawFireSquare extSquare
End Sub

