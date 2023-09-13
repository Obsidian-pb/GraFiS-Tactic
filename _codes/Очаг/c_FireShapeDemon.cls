VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireShapeDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private grain As Integer                            'Размер зерна
Private fireShp As Visio.Shape                      'Анализируемая фигура


Const mmInInch As Double = 25.4

'--------------------------------Демон построитель зоны горения-------------------------------------------
'1 - Получает на вход фигуру площади горения и анализирует ее, заполняя fireMatrixObject соответствующими значениями


Public Function setGrain(ByVal gr As Integer) As c_FireShapeDemon
    grain = gr
    Set setGrain = Me
End Function

Public Function SetFireShape(ByRef shp As Visio.Shape) As c_FireShapeDemon
    Set fireShp = shp
    Set SetFireShape = Me
End Function



Public Sub RunDemon(ByRef fireMatrixObject As c_Matrix)
'Запускаем демона
'Основаня функция демона - анализ фигуры
Dim offsetX As Double                           'Смещение клеток по оси X, мм
Dim offsetY As Double                           'Смещение клеток по оси Y, мм
Dim offsetXC As Double                          'Смещение клеток по оси X, клетки
Dim offsetYC As Double                          'Смещение клеток по оси Y, клетки
Dim sizeXC As Long
Dim sizeYC As Long
Dim i As Long
Dim j As Long
Dim x As Double
Dim y As Double
Dim cell As c_Cell
Dim shp As Visio.Shape

    '1 - опредеяем смещения клеток
    offsetX = fireShp.Cells("PinX").Result(visMillimeters) - fireShp.Cells("LocPinX").Result(visMillimeters)
    offsetY = fireShp.Cells("PinY").Result(visMillimeters) - fireShp.Cells("LocPinY").Result(visMillimeters)
    offsetXC = offsetX / grain
    offsetYC = offsetY / grain
    
    '2 - определяем размер контура фигуры (чтобы не считать вообще все на листе)
    sizeXC = Int(fireShp.Cells("Width").Result(visMillimeters) / grain)
    sizeYC = Int(fireShp.Cells("Height").Result(visMillimeters) / grain)
    
    '3 - перебираем все клетки и проверяем присутствует ли в них фигура
    For i = 0 To sizeXC
        For j = 0 To sizeYC
            x = (i * grain + offsetX + 0.5 * grain) / mmInInch
            y = (j * grain + offsetY + 0.5 * grain) / mmInInch
            
            If fireShp.HitTest(x, y, 0) > 0 Then
'                fireMatrixObject.SetCellValue offsetXC + i + 1, offsetYC + j + 1, mtrCurrentgPowerLayer, 100 ' CellState.csFire
'                fireMatrixObject.SetFiredCell offsetXC + i + 1, offsetYC + j + 1
'                if fireMatrixObject.is
'                fireMatrixObject.SetCellValue offsetXC + i, offsetYC + j, mtrCurrentgPowerLayer, 100   ' CellState.csFire
                fireMatrixObject.SetFiredCell offsetXC + i, offsetYC + j, 100
                
                
'                Set shp = Application.ActivePage.DrawRectangle(((offsetXC + i) * grain) / mmInInch, ((offsetYC + j) * grain) / mmInInch, _
'                                                     ((offsetXC + i) * grain + grain) / mmInInch, ((offsetYC + j) * grain + grain) / mmInInch)
'                                                     shp.Cells("FillForegnd").Formula = 2
'                                                     shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("FireSpace")
            End If
        Next j
    Next i
    
    '4 - Определяем клетки фронта и периметра фигуры
    For Each cell In fireMatrixObject.firedCells
        'Фронт пожара
        If Not fireMatrixObject.IsInner(cell.x, cell.y) Then
            AddToFrontCells fireMatrixObject, cell          'Добавляем клетку в коллекцию клеток фронта пожара

'                Set shp = Application.ActivePage.DrawRectangle((cell.x * grain) / mmInInch, (cell.y * grain) / mmInInch, _
'                                                     (cell.x * grain + grain) / mmInInch, (cell.y * grain + grain) / mmInInch)
'                                                     shp.Cells("LinePattern").Formula = 0
'                                                     shp.Cells("FillForegnd").Formula = 2
'                                                     shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = GetLayerNumber("FireFront")
        End If
        'Периметр пожара
        If Not fireMatrixObject.IsNotPerimeter(cell.x, cell.y) Then
            AddToPerimeterCells fireMatrixObject, cell          'Добавляем клетку в коллекцию клеток фронта пожара
        End If
    Next cell
    
End Sub


Private Sub AddToFrontCells(ByRef fireMatrixObject As c_Matrix, ByRef cell As c_Cell)
'Добавляем клетку в коллекцию клеток фронта пожара. Если же клетка с указанным ключем уже имеется - выходим
    On Error Resume Next
    fireMatrixObject.fireFrontCells.Add cell, cell.x & ":" & cell.y
End Sub
Private Sub AddToPerimeterCells(ByRef fireMatrixObject As c_Matrix, ByRef cell As c_Cell)
'Добавляем клетку в коллекцию клеток периметра. Если же клетка с указанным ключем уже имеется - выходим
    On Error Resume Next
    fireMatrixObject.firePerimeterCells.Add cell, cell.x & ":" & cell.y
End Sub
