Attribute VB_Name = "t_GFSTools"
Public Function IsGFSShape(ByRef shp As Visio.Shape, Optional ByVal useManeure As Boolean = True) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС
Dim i As Integer
    
'    If shp.CellExists("User.IndexPers", 0) = True and shp.CellExists("User.Version", 0) = True Then        'Подумать - нужен ли вообще учет версии
    'Проверяем, является ли фигура фигурой ГраФиС
    If useManeure Then      'Если нужно учитывать проверку на маневр
        If shp.CellExists("User.IndexPers", 0) = True Then
            'Если имеется ячейка опции Маневра и ее значение показывает, что
            If shp.CellExists("Actions.MainManeure", 0) = True Then
                If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
                    IsGFSShape = True       'Фигура ГраФиС и не маневренная
                Else
                    IsGFSShape = False      'Фигура ГраФиС и маневренная
                End If
            Else
                IsGFSShape = True       'Фигура ГраФиС и не имеет ячейки Маневр
            End If
        Else
            IsGFSShape = False      'Фигура не ГраФиС
        End If
    Else                    'если не нужно учитывать проверку на маневр
'        If shp.CellExists("User.IndexPers", 0) = True Then
'            IsGFSShape = True       'Фигура ГраФиС
'        Else
'            IsGFSShape = False      'Фигура не ГраФиС
'        End If
        IsGFSShape = shp.CellExists("User.IndexPers", 0)
    End If

End Function

Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPerses As Variant, Optional needGFSChecj As Boolean = False) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС и среди переданных типов фигур ГраФиС (gfsIndexPreses) присутствует IndexPers данной фигуры
'По умолчанию предполагается что переданная фигура уже проверена на то, относится ли она к фигурам ГраФиС. В случае, если у фигуры нет ячейки User.IndexPers _
'обработчик ошибки указывает функции вернуть False
'Пример использования: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
'                 или: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
Dim i As Integer
Dim indexPers As Integer
    
    On Error GoTo EX
    
    'Если необходима предварительная проверка на отношение фигуры к ГраФиС:
    If needGFSChecj Then
        If Not IsGFSShape(shp) Then
            IsGFSShapeWithIP = False
            Exit Function
        End If
    End If
    
    'Проверяем, является ли фигура фигурой указанного типа
    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
    Select Case TypeName(gfsIndexPerses)
        Case Is = "Long"    'Если передано единственное значение Long
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Integer"    'Если передано единственное значение Integer
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Variant()"   'Если передан массив
            For i = 0 To UBound(gfsIndexPerses)
                If gfsIndexPerses(i) = indexPers Then
                    IsGFSShapeWithIP = True
                    Exit Function
                End If
            Next i
        Case Else
            IsGFSShapeWithIP = False
    End Select

IsGFSShapeWithIP = False
Exit Function
EX:
    IsGFSShapeWithIP = False
    SaveLog Err, "m_Tools.IsGFSShapeWithIP"
End Function

