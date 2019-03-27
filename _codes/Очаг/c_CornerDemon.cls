VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CornerDemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------Класс демона движущегося вдоль стен------------------------------------------------
Const PI = 3.1415926535

Private x As Long
Private y As Long
Private startX As Long
Private startY As Long
Private direction As Directions

Private mainAngle As Double
Private perpAngle As Double

Private matrix As c_Matrix
Public lineFillDemon As c_CornerFillDemon




'-----------------Процедуры работы с классом извне----------------------------------------
Public Sub Init(ByRef matrixRef As c_Matrix, ByVal xVal As Long, ByVal yVal As Long)
'Инициализация класса
    direction = s       'Стоит на месте
    
    'Передаем матрицу
    Set matrix = matrixRef
    
    'Активируем демона-отрисовщика линий
    Set lineFillDemon = New c_CornerFillDemon
    lineFillDemon.Init matrix
    
    'Задаем стартовые позиции
    SetStartPosition xVal, yVal
End Sub

Public Sub SetStartPosition(ByVal xVal As Long, ByVal yVal As Long)
'Ставим стартовые координаты демона
    startX = xVal
    startY = yVal
    SetPosition xVal, yVal
End Sub

Private Sub SetPosition(ByVal xVal As Long, ByVal yVal As Long)
'Ставим координаты демона
    x = xVal
    y = yVal
End Sub

Public Sub SetDirection(ByVal dirct As Directions)
'Задаем направление движения демона
    direction = dirct
End Sub

Public Sub RunDemon()
'Запукаем демона
Dim i As Integer
    
    'Запускаем другого демона для определения угла перпендикулярного стене
    Dim angleDemon As c_AngleDemon
    Set angleDemon = New c_AngleDemon
    angleDemon.Init matrix, x, y, direction
    angleDemon.RunDemon
    
    mainAngle = angleDemon.mainAngle
    perpAngle = angleDemon.perpAngle
    
    
    
    'Демон пробегает по всем клеткам вдоль стены и каждый раз проверяет перпендикуляр к ней. Если демон не может корректно продолжить движение, либо же демон заливки не нашел по перпендикуляру горящих клеток, заканчиваем задание
    i = 0
    Do While NextCell
        i = i + 1
        'Проверяем не выходит ли данная клетка за пределы допустимого угла
        If Not isPossiblePlace Then Exit Do
        
        'Запускаем демона закрашивающего линию от текущего положения
        If lineFillDemon.testLine(x, y, perpAngle) = False Then Exit Do
        
        'Байпасс - экстренный выход, если что-то пошло не так
        If i > 100 Then
            Exit Do
        End If
    Loop
       
End Sub




'-------------------------------------------Внутренние процедуры класса----------------------------------------
Private Function NextCell() As Boolean
'Выбираем следующую клетку

    'Если демон пока не движется - ищем направление
    If direction = s Then
        direction = FindDirection
        
        'Если найти направление движения не удалось - выходим
        If direction = s Then
            NextCell = False
            Exit Function
        End If
        
        'Перемещаем демона
        MoveDemon
        NextCell = True
        Exit Function
    Else
        'Ищем клетки в направлении прямо по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(0)) Then
            direction = NewDirection(0)
            'Перемещаем демона
            MoveDemon
            NextCell = True
            Exit Function
        End If
        'Ищем клетки в направлении левее по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(-2)) Then
            direction = NewDirection(-2)
            'Перемещаем демона
            MoveDemon
            NextCell = True
            Exit Function
        End If
        'Ищем клетки в направлении правее по пути следования
        If IsCellNearTheWallAndNotFireByDirection(NewDirection(2)) Then
            direction = NewDirection(2)
            'Перемещаем демона
            MoveDemon
            NextCell = True
            Exit Function
        End If
    End If
    
    'Дальнейший путь не найден
    direction = s
    MoveDemon
    NextCell = False
    
End Function

Private Sub MoveDemon()
'Перемещаем демона в текущем направлении
    Select Case direction
        Case Is = l
            x = x - 1
        Case Is = lu
            y = y + 1
            x = x - 1
        Case Is = u
            y = y + 1
        Case Is = ru
            y = y + 1
            x = x + 1
        Case Is = r
            x = x + 1
        Case Is = rd
            y = y - 1
            x = x + 1
        Case Is = d
            y = y - 1
        Case Is = ld
            y = y - 1
            x = x - 1
        Case Else
'            Debug.Print "Стоит!!! x=" & x & ", y=" & y
    End Select
End Sub

Private Function IsCellNearTheWallAndNotFireByDirection(ByVal direct As Directions) As Boolean
    Select Case direct
        Case Is = l
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y)
        Case Is = lu
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y + 1)
        Case Is = u
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y + 1)
        Case Is = ru
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y + 1)
        Case Is = r
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y)
        Case Is = rd
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x + 1, y - 1)
        Case Is = d
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x, y - 1)
        Case Is = ld
            IsCellNearTheWallAndNotFireByDirection = IsCellNearTheWallAndNotFire(x - 1, y - 1)
    End Select
End Function

Private Function IsCellNearTheWallAndNotFire(ByVal xVal As Long, ByVal yVal As Long) As Boolean
'Провеяем находится ли клетка рядом со стеной.  И при этом не полностью горит и сама не является стеной
    'Проверяем, не выходит ли координата ячейки за пределы матрицы
    If x < 0 Or y < 0 Or x > matrix.DimensionX Or y > matrix.DimensionY Then Exit Function
    
    'Проверяем саму клетку
    If matrix.IsCellMaxFire(xVal, yVal) Or matrix.IsWall(xVal, yVal) Then
        IsCellNearTheWallAndNotFire = False
        Exit Function
    End If
    
    'По ортогонали
    If matrix.IsWall(xVal, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal - 1, yVal) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal + 1, yVal) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    'По диагонали
    If matrix.IsWall(xVal - 1, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal - 1, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal + 1, yVal - 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    If matrix.IsWall(xVal + 1, yVal + 1) Then
        IsCellNearTheWallAndNotFire = True
        Exit Function
    End If
    
    
IsCellNearTheWallAndNotFire = False
End Function

Private Function NewDirection(ByVal directionDifferent As Integer) As Directions
'Получаем смещение направления в зависимости о тпереданного дифференциала
Dim tmpVal As Integer
    
    tmpVal = direction + directionDifferent
    
    If tmpVal < 1 Then tmpVal = tmpVal + 8
    If tmpVal > 8 Then tmpVal = tmpVal - 8
    
NewDirection = tmpVal
End Function

Private Function FindDirection() As Directions
    
    'по диагонали
    If IsCellNearTheWallAndNotFire(x - 1, y - 1) Then
        FindDirection = ld
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y - 1) Then
        FindDirection = rd
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y + 1) Then
        FindDirection = lu
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y + 1) Then
        FindDirection = ru
        Exit Function
    End If
    'по ортогонали
    If IsCellNearTheWallAndNotFire(x, y - 1) Then
        FindDirection = d
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x, y + 1) Then
        FindDirection = u
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x - 1, y) Then
        FindDirection = l
        Exit Function
    End If
    If IsCellNearTheWallAndNotFire(x + 1, y) Then
        FindDirection = r
        Exit Function
    End If
End Function

Private Function GetAngle() As Double
'Возвращает угол между стартовой позицией и текущим положением демона
Dim xDiff As Long
Dim yDiff As Long

    On Error Resume Next

    xDiff = x - startX
    yDiff = y - startY


    If xDiff = 0 Then
        If y > startY Then
            GetAngle = 90
        Else
            GetAngle = -90
        End If
    Else
        'Обрабатываем четверти
        '1-я четверть
        If xDiff > 0 And yDiff > 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI
        End If
        '2-я четверть
        If xDiff < 0 And yDiff > 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI + 180
        End If
        '3-я четверть
        If xDiff < 0 And yDiff < 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI + 180
        End If
        '4-я четверть
        If xDiff > 0 And yDiff < 0 Then
            GetAngle = Atn(yDiff / xDiff) * 180 / PI
        End If
        
    End If

End Function

Private Function isPossiblePlace() As Boolean
'Функция возвращает Истина, если данные координаты допустимы для клетки в рамках линии с данным углом
Dim possibleDiff As Double
Dim tolerance As Single
Dim dist As Double
Dim currentAngle As Double
    
    'Определяем допустимый предел отклонения угла
    dist = Sqr((x - startX) ^ 2 + (y - startY) ^ 2)
    If dist = 0 Then
        isPossiblePlace = True
        Exit Function
    End If
    
    tolerance = 1.01
    possibleDiff = Atn(1 / dist) * (180 / PI) * tolerance
    
    'Получаем текущий угол
    currentAngle = GetAngle
    
    'Проверяем, лежит ли угол к текущей клетке в допустимых пределах
    If currentAngle < mainAngle + possibleDiff And currentAngle > mainAngle - possibleDiff Then
        isPossiblePlace = True
    Else
        isPossiblePlace = False
    End If
    
End Function
