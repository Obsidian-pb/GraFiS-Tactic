Attribute VB_Name = "m_Tools"
'-----------------------------------------Модуль инструментальных функций----------------------------------------------
Option Explicit



Public Function PF_RoundUp(afs_Value As Single) As Integer
'Процедура округления ПОЛОЖИТЕЛЬНЫХ чисел в большую сторону
Dim vfi_Temp As Integer

vfi_Temp = Int(afs_Value * (-1)) * (-1)
PF_RoundUp = vfi_Temp

End Function

Public Function CellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber) As Variant
'Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
    
    On Error GoTo EX
    
    Select Case dataType
        Case Is = visNumber
            CellVal = shp.Cells(cellName).Result(dataType)
        Case Is = visUnitsString
            CellVal = shp.Cells(cellName).ResultStr(dataType)
    End Select
    
    
    
Exit Function
EX:
    CellVal = 0
End Function

Public Function IsGFSShape(ByRef shp As Visio.Shape) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС
Dim i As Integer
    
    'Проверяем, является ли фигура фигурой ГраФиС
    If shp.CellExists("User.IndexPers", 0) = True Then
        IsGFSShape = True
        Exit Function
    End If
    
IsGFSShape = False
End Function

Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPreses As Variant) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС и среди переданных типов фигур ГраФиС (gfsIndexPreses) присутствует IndexPers данной фигуры
'Предполагается что переданная фигура уже проверена на то, относится ли она уже к фигурам ГраФиС. В случае, если у фигуры нет ячейки User.IndexPers _
'обработчик ошибки указывает функции вернуть False
'Пример использования: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
'                 или: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
Dim i As Integer
Dim indexPers As Integer
    
    On Error GoTo EX
    
    'Проверяем, является ли фигура фигурой указанного типа
    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
    Select Case TypeName(gfsIndexPreses)
        Case Is = "Long"    'Если передано единственное значение
            If gfsIndexPreses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Variant()"   'Если передан массив
            For i = 0 To UBound(gfsIndexPreses)
                If gfsIndexPreses(i) = indexPers Then
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


