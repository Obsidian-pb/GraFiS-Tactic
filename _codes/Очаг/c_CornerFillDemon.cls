VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CornerFillDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------Класс демона заливающего промежутки между стенами и зонами горения (переход к прямоугольной форме)----------------------------
Const PI = 3.1415926535

Private matrix As c_Matrix

Public cellsForFill As Collection







'------------------------Публичные процедуры класса---------------------------------
Public Sub Init(ByRef matrixRef As c_Matrix)
'Инициализируем класс
    Set matrix = matrixRef
    Set cellsForFill = New Collection
End Sub

Public Function testLine(ByVal startX As Long, ByVal startY As Long, ByVal angle As Double)
'Проверяем, упирается ли линия проведенная из определенной точки с определенным углом в горящую зону
Dim x As Integer
Dim y As Integer
Dim cs As Double
Dim sn As Double
Dim i As Integer
Dim j As Integer
Dim CellType As Byte
    
    'Добавляем стартовую клетку
'    AddToCellCollection startX, startY
    
    'Заранее определяем косинус и синус угла - для дальнейшего построения линии
    cs = Cos(angle * PI / 180)
    sn = Sin(angle * PI / 180)
    
    'Последовательно проверяем все клеткии линии, на предмет того, являются они пустой клеткой, горящей клеткой или же клеткой стены
    i = 0
    Do While CellType = 0
        x = startX + Int(cs * i)
        y = startY + Int(sn * i)
        
        CellType = GetKindOfCell(x, y)
        
        i = i + 1
        If i > 100 Then Exit Do
    Loop

    
    'Закрашиваем линию, если она уперлась в зону горения
'    '!!!Временно, для проверки
'    If cellType = 1 Then
'        testLine = True
'        Exit Function
'    End If

    'Вносим все клетки в коллекцию клеток для последующего закрашивания
    If CellType = 1 Then
        For j = 0 To i - 2
            x = startX + Int(cs * j)
            y = startY + Int(sn * j)
            
            AddToCellCollection x, y
        Next j
    
        'Возвращаем Истина, если демон наткнулся на горящую клетку
        testLine = True
        Exit Function
    End If
    
    'Возвращаем Ложь, если демон НЕ наткнулся на горящую клетку
    testLine = False
End Function


Private Function GetKindOfCell(ByVal xVal As Long, ByVal yVal As Long) As Byte
'Возвращает тип клетки (0-свободная, 1-горит на максимум, 2-непроходима-прекратить проверку)
    On Error GoTo EX
    If matrix.IsCellMaxFire(xVal, yVal) Then
        GetKindOfCell = 1
        Exit Function
    End If
    If matrix.IsWall(xVal, yVal) Then
        GetKindOfCell = 2
        Exit Function
    End If
    GetKindOfCell = 0
Exit Function
EX:
    GetKindOfCell = 2
End Function


Private Sub AddToCellCollection(ByVal xVal As Long, ByVal yVal As Long)
'Добавляем клетку в коллекцию
    Dim cell As c_Cell
    Set cell = New c_Cell
    cell.x = xVal
    cell.y = yVal
    
    On Error Resume Next
    cellsForFill.Add cell, xVal & ":" & yVal
End Sub


'Public Sub FillCells()
''Заливаем все клетки из коллекции
'Dim cell As c_Cell
'
'    For Each cell In cellsForFill
'        Cells(cell.y, cell.x) = 100
'    Next cell
'End Sub













