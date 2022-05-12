Attribute VB_Name = "m_MatrixWork"
Option Explicit
'--------------------------------------Модуль импорта/экспорта матриц----------------------





Public Sub SaveMatrixTo(Optional path As String = "")
'Сохраняем матрицу в csv файл
Dim mLayer As Variant
Dim layerString As String
    
    '---Печатаем размер массива слоя открытых пространств
    arrShape fireModeller.GetOpenSpaceLayer
    
    '---Получаем массив слоя открытых пространств матрицы и преобразуем его в строку
    layerString = Array2DToString(fireModeller.GetOpenSpaceLayer)
    
    '---Если путь к файлу не был передан, указываем путь по-умолчанию для документа
    If path = "" Then
        path = Replace(Application.ActiveDocument.fullName, ".vsdx", ".csv")
        path = Replace(path, ".vsd", ".csv")
    End If
    
    '---Сохраняем в файл
    SaveTextToFile layerString, path
    
    '---Печатаем в дебаг размер массива
     Debug.Print arrShape(fireModeller.GetOpenSpaceLayer)
End Sub

Private Function Array2DToString(arr As Variant) As String
Dim i As Integer
Dim j As Integer
Dim s As String
    

    For j = 0 To UBound(arr, 2)
        For i = 0 To UBound(arr, 1)
            s = s + CStr(arr(i, j)) & ","
        Next i
    Next j

    
Array2DToString = Left(s, Len(s) - 1)
End Function


Public Function arrShape(arr As Variant) As String
'Получаем размерность массива
Dim i As Integer
Dim s As String
    
    On Error Resume Next
    
    For i = 1 To 10
        s = s & UBound(arr, i) + 1 & ","
    Next i

arrShape = Left(s, Len(s) - 1)
End Function



'тесты
Public Sub AAA()
Dim a(1, 2) As Integer
    
    a(0, 0) = 10
    a(0, 1) = 20
    a(0, 2) = 30
    a(1, 0) = 40
    a(1, 1) = 50
    a(1, 2) = 60
    
    Debug.Print arrShape(a)
    Debug.Print Array2DToString(a)
End Sub

