Attribute VB_Name = "m_MatrixImportExport"
Option Explicit


'---------------------------Модуль для процедур Импорта/Экспорта----------------------------------
Public Sub SaveMatrixTo(Optional path As String = "")
'Сохраняе матрицу открытых пространств в формате массива numpy в csv файл
Dim lay As Variant
Dim s As String
Dim x As Integer
Dim y As Integer

    If path = "" Then
        path = Replace(Application.ActiveDocument.fullName, ".vsdx", ".csv")
        path = Replace(path, ".vsd", ".csv")
    End If
    
'---Получаем матрицу открытых пространств
    lay = fireModeller.GetOpenSpaceLayer

'---Открываем файл матрицы в csv (если его нет - создаем)
    Open path For Output As #1
    
'---Формируем строку матрицы
    For y = 0 To UBound(lay, 2)
        s = ""
        For x = 0 To UBound(lay, 1)
            s = s & CStr(lay(x, y)) & ","
        Next x
    '---Записываем в конец файла лога сведения о ошибке
        Print #1, Left(s, Len(s) - 1)
    Next y

'---Закрываем файл лога
    Close #1
End Sub



