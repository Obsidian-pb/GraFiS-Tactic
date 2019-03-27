VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FireBuilderT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''----------------Класс для отрисовки плоащди пожара согласно тактического метода построения------------------------------------------
'Option Explicit
'Option Base 0
'
'
'Public grain As Integer
'Public way As Double
'
'
'Private CirclseStampCol As Collection
'Private CircleParts(7) As Collection
'
'
'Private perimeterPointsCol As Collection
'Private perimeterPointsColTemp As Collection
'Private perimeterPointsArrayString As String
'
'Const circleRadiuss As Integer = 15     'в зернах   15
'
'
'
'
'Public Sub Init(ByVal grainVal As Integer, ByVal wayVal As Double)
''Задаем стартовые значения переменных
'    grain = grainVal
'    way = wayVal
'
'    '---Объявляем коллекцию-штампы секторов круга который будет использоваться для отрисовки
'    Dim i As Integer
'    For i = 0 To 7
'        Set CircleParts(i) = New Collection
'    Next i
'    '---Формируем коллекцию-штамп круга который будет использоваться для отрисовки и его секторов
'    GetCircleStampCollection
'
'End Sub
'
'
'
'
'
'
'Public Sub BuildFire(ByRef matrix() As Byte, ByRef currentFirePointCol As Collection, ByVal rounds As Integer)
''Стартовый метод построения площади
'Dim firePoint As c_Point
'Dim circlePoint As c_Point
'Dim newPoint As c_Point
'Dim currentCollection As Collection
'Dim newPoints As Collection
'Dim i As Integer
'
'Dim pnt As c_Point
'Dim pnt2 As c_Point
'
'    '---Строим пожар от стартовых точек
'    SetCollection perimeterPointsCol, currentFirePointCol
'
'
'    For i = 0 To rounds - 1
'        SetCollection currentCollection, perimeterPointsCol
'
'        '---Очищаем коллекцию периметра от флага внешней зоны пожара
'        ClearPerimeterFromOuterFlag matrix, perimeterPointsCol
'
'        Set newPoints = New Collection
'        '---Строим новые зоны горения
'        For Each firePoint In currentCollection
'            MakeShot matrix, firePoint, newPoints ' currentFirePointCol
'        Next firePoint
'
'        '---отмечаем новые клетки как горящие
'        For Each firePoint In newPoints
'            matrix(firePoint.x, firePoint.y) = CellState.csFire
'        Next firePoint
'
'        AddCollectionItems currentFirePointCol, newPoints
'
'        '---Получаем периметр
'        Set perimeterPointsCol = GetPerimeterPointsCol(matrix, currentFirePointCol, False)
'
'
'
'
''        For Each pnt In perimeterPointsCol
''            Application.ActivePage.DrawRectangle pnt.x * grain, pnt.y * grain, _
''                            (pnt.x * grain + grain), (pnt.y * grain + grain)
''        Next pnt
'
'        '---Нормализуем периметр
'        Set perimeterPointsColTemp = New Collection
'        GetPerimeterPointsCollection matrix, perimeterPointsCol.item(1)
'        NormalizeCollection perimeterPointsColTemp
'
'        '---Рисуем кривую
'        Application.ActivePage.DrawPolyline GetPerimeterPointsDoubleArray(perimeterPointsColTemp), visPolyline1D
'        DoEvents
'
'    Next i
'
''        For Each pnt In currentFirePointCol
''            Application.ActivePage.DrawRectangle pnt.x * grain, pnt.y * grain, _
''                            (pnt.x * grain + grain), (pnt.y * grain + grain)
''        Next pnt
'
''    '---Рисуем периметр
''    Set perimeterPointsColTemp = New Collection
''    GetPerimeterPointsCollection matrix, perimeterPointsCol.item(1)
''
''    '---Нормализуем коллекцию координат для отрисовки
''    NormalizeCollection perimeterPointsColTemp
''
''    '---Рисуем кривую
''    Application.ActivePage.DrawPolyline GetPerimeterPointsDoubleArray(perimeterPointsColTemp), visPolyline1D
'
'End Sub
'
'
'Private Sub MakeShot(ByRef matrix() As Byte, ByRef point As c_Point, ByRef currentFirePointCol As Collection)
'
'Dim pnt As c_Point
'
'    If matrix(point.x - 1, point.y) = CellState.csOpenSpace Or matrix(point.x - 1, point.y) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(0)
'    End If
'    If matrix(point.x - 1, point.y + 1) = CellState.csOpenSpace Or matrix(point.x - 1, point.y + 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(1)
'    End If
'    If matrix(point.x, point.y + 1) = CellState.csOpenSpace Or matrix(point.x, point.y + 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(2)
'    End If
'    If matrix(point.x + 1, point.y + 1) = CellState.csOpenSpace Or matrix(point.x + 1, point.y + 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(3)
'    End If
'    If matrix(point.x + 1, point.y) = CellState.csOpenSpace Or matrix(point.x + 1, point.y) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(4)
'    End If
'    If matrix(point.x + 1, point.y - 1) = CellState.csOpenSpace Or matrix(point.x + 1, point.y - 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(5)
'    End If
'    If matrix(point.x, point.y - 1) = CellState.csOpenSpace Or matrix(point.x, point.y - 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(6)
'    End If
'    If matrix(point.x - 1, point.y - 1) = CellState.csOpenSpace Or matrix(point.x - 1, point.y - 1) = CellState.csWillBurnNextStep Then
'        MakeSector matrix, NewCPoint(point.x, point.y), currentFirePointCol, CircleParts(7)
'    End If
'
'
'
'End Sub
'
'Private Function NewCPoint(ByVal x As Double, ByVal y As Double) As c_Point
'Dim newPoint As c_Point
'    Set newPoint = New c_Point
'    newPoint.SetData x, y
'    Set NewCPoint = newPoint
'End Function
'
'Private Sub MakeSector(ByRef matrix() As Byte, ByRef point As c_Point, ByRef currentFirePointCol As Collection, ByRef sectorPoints As Collection)
'
'Dim sectorPoint As c_Point
'
'    For Each sectorPoint In sectorPoints
'        If matrix(point.x + sectorPoint.x, point.y + sectorPoint.y) = CellState.csOpenSpace Then
'            matrix(point.x + sectorPoint.x, point.y + sectorPoint.y) = CellState.csWillBurnNextStep
'            currentFirePointCol.Add NewCPoint(point.x + sectorPoint.x, point.y + sectorPoint.y)
'
''            Application.ActivePage.DrawRectangle (point.x + sectorPoint.x) * grain, (point.y + sectorPoint.y) * grain, _
''                            (point.x + sectorPoint.x) * grain + grain, (point.y + sectorPoint.y) * grain + grain
''            DoEvents
'        End If
'    Next sectorPoint
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'
'Private Sub GetCircleStampCollection()
''Возвращает коллекцию точек в штампе круга с заданным диаметром
'Dim i As Integer
'Dim j As Integer
'Dim pnt As c_Point
'Dim angle As Double
'
'    Set CirclseStampCol = New Collection
'
'
'    For i = -circleRadiuss To circleRadiuss
'        For j = -circleRadiuss To circleRadiuss
'            If SpecialRound(Sqr(i ^ 2 + j ^ 2), 0.25) <= circleRadiuss Then
'                Set pnt = New c_Point
'                pnt.SetData i, j
'
'                CirclseStampCol.Add pnt
'
'                '---Разбираем точки по коллекциям сегментов     atn(y/x)* (180/3.14)
'                angle = pnt.ATANCorr
'                '0
'                If angle >= 157.5 And angle < 202.5 Then
'                    CircleParts(0).Add pnt
'                End If
'                '1
'                If angle >= 112.5 And angle < 157.5 Then
'                    CircleParts(1).Add pnt
'                End If
'                '2
'                If angle >= 67.5 And angle < 112.5 Then
'                    CircleParts(2).Add pnt
'                End If
'                '3
'                If angle >= 22.5 And angle < 67.5 Then
'                    CircleParts(3).Add pnt
'                End If
'                '4
'                If angle >= 337.5 Or angle < 22.5 Then
'                    CircleParts(4).Add pnt
'                End If
'                '5
'                If angle >= 292.5 And angle < 337.5 Then
'                    CircleParts(5).Add pnt
'                End If
'                '6
'                If angle >= 247.5 And angle < 292.5 Then
'                    CircleParts(6).Add pnt
'                End If
'                '7
'                If angle >= 202.5 And angle < 247.5 Then
'                    CircleParts(7).Add pnt
'                End If
'            End If
'
'        Next j
'    Next i
'
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'
'
'Private Function GetPerimeterPointsCol(ByRef matrix() As Byte, ByRef currentFirePointCol As Collection, _
'                                        Optional ByVal diagonals As Boolean) As Collection
''Возвращает коллекцию точек периметра
'Dim pnt As c_Point
'Dim tmpCol As Collection
'
'    Set tmpCol = New Collection
'
'    For Each pnt In currentFirePointCol
'        If IsOuter(matrix, pnt, diagonals) Then
'            tmpCol.Add pnt
'        End If
'    Next pnt
'
'Set GetPerimeterPointsCol = tmpCol
'End Function
'
'Private Function IsOuter(ByRef matrix() As Byte, ByRef point As c_Point, Optional ByVal diagonals As Boolean) As Boolean
''опрелделяем, является ли указанная клетка внешней
'    If diagonals Then
'        If matrix(point.x - 1, point.y - 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x + 1, point.y + 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x + 1, point.y - 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x - 1, point.y + 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'    Else
'        If matrix(point.x - 1, point.y) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x + 1, point.y) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x, point.y - 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'        If matrix(point.x, point.y + 1) = CellState.csOpenSpace Then
'            IsOuter = True
'            matrix(point.x, point.y) = CellState.csFireOuter
'            Exit Function
'        End If
'    End If
'
'
'IsOuter = False
'End Function
'
'Private Function GetPerimeterPointsDoubleArray(ByRef coordsCollection As Collection) As Double()
''Получаем цепочку координат точек периметра в виде массива Double
'Dim pnt As c_Point
'Dim arr() As Double
'Dim i As Long
'
'    ReDim arr(coordsCollection.Count * 2 + 1)
'
'    i = 0
'    For Each pnt In coordsCollection
'        arr(i) = pnt.x * grain / 25.4
'        arr(i + 1) = pnt.y * grain / 25.4
'        i = i + 2
'    Next pnt
'
'    arr(coordsCollection.Count * 2) = arr(0)
'    arr(coordsCollection.Count * 2 + 1) = arr(1)
'
'GetPerimeterPointsDoubleArray = arr
'End Function
'
'Private Function GetPerimeterPointsCollection(ByRef matrix() As Byte, ByRef point As c_Point) As Boolean
''Получаем цепочку координат точек периметра в виде коллекции
'Dim i As Integer
'Dim j As Integer
'Dim newPoint As c_Point
'
'    If matrix(point.x, point.y) = CellState.csFireOuter And Not IsInCollection(perimeterPointsColTemp, point) Then
'        perimeterPointsColTemp.Add point
'
'        For i = -1 To 1
'            For j = -1 To 1
'                GetPerimeterPointsCollection matrix, NewCPoint(point.x + i, point.y + j)
'            Next j
'        Next i
'
'        GetPerimeterPointsCollection = True
'        Exit Function
'    End If
'
'
'GetPerimeterPointsCollection = False
'End Function
'
'Private Sub ClearPerimeterFromOuterFlag(ByRef matrix() As Byte, ByRef perimeterPointsCollection As Collection)
''Очищаем коллекцию точек периметра от пометки, что они внешние
'Dim pnt As c_Point
'
'    For Each pnt In perimeterPointsCollection
'        matrix(pnt.x, pnt.y) = CellState.csFire
'    Next
'End Sub
'
'
'
