Attribute VB_Name = "m_Tools"
Option Explicit





Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 44 Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function

Public Function PFB_isPlace(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - место, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isPlace = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 5 And aO_Shape.Cells("User.ShapeType").Result(visNumber) = 38 Then
        PFB_isPlace = True
        Exit Function
    End If
PFB_isPlace = False
End Function

Public Function PFI_FirstSectionCount(ByRef aO_Shape As Visio.Shape) As Integer
'Функция возвращает количество графических секций
Dim i As Integer

    i = 0
    Do While aO_Shape.SectionExists(visSectionFirstComponent + i, 0)
        i = i + 1
    Loop
    
PFI_FirstSectionCount = i
End Function

Public Function PF_DocumentOpened(ByVal DocName As String) As Boolean
'Функция возвращает Истина, если документ открыт, Ложь, если нет
Dim vO_Doc As Visio.Document

    For Each vO_Doc In Application.Documents
        If InStr(1, vO_Doc.Name, DocName, vbTextCompare) Then
            PF_DocumentOpened = True
            Exit Function
        End If
    Next vO_Doc
PF_DocumentOpened = False
End Function

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

'-----------------------------------------Процедуры работы с фигурами----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'Процедура устанавливает новое значение для всех выбранных фигур одного типа
Dim shp As Visio.Shape
    
    'Перебираем все фигуры в выделении и если очередная фигура имеет такую же ячейку - присваиваем ей новое значение
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).Formula = aB_Value
        End If
    Next shp
    
End Sub

'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub


