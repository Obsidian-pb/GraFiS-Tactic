Attribute VB_Name = "m_MatrixImportExport"
Option Explicit


'---------------------------Модуль для процедур Импорта/Экспорта----------------------------------
Public Sub SaveMatrixTo(Optional path As String = "")
'Сохраняе матрицу открытых пространств в формате массива numpy в csv файл
Dim lay As Variant
Dim s As String
Dim X As Integer
Dim Y As Integer

    If path = "" Then
        path = Replace(Application.ActiveDocument.fullName, ".vsdx", ".csv")
        path = Replace(path, ".vsd", ".csv")
    End If

'---Получаем матрицу открытых пространств
    lay = fireModeller.GetOpenSpaceLayer

'---Открываем файл матрицы в csv (если его нет - создаем)
    Open path For Output As #1
    
'---Формируем строку матрицы
    For Y = 0 To UBound(lay, 2)
        s = ""
        For X = 0 To UBound(lay, 1)
            If IsEmpty(lay(X, Y)) Then
                s = s & "2,"
            Else
                s = s & CStr(lay(X, Y)) & ","
            End If
        Next X
    '---Записываем в конец файла лога сведения о ошибке
        Print #1, Left(s, Len(s) - 1)
    Next Y

'---Закрываем файл лога
    Close #1
End Sub



