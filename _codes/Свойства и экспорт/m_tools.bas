Attribute VB_Name = "m_tools"
Option Explicit



Public Function GetReadyString(ByVal val As Variant, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'Возвращает сформированную строку в виде prefix & val & postfix. В случае если val=ignore, возвращает ifEmpty
'В случае невозможности преобразования val к строке возвращает ifEmpty
    
    On Error GoTo ex
    
    If val = ignore Then
        GetReadyString = ifEmpty
    Else
        GetReadyString = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyString = ifEmpty
End Function
Public Function GetReadyStringA(ByVal elemID As String, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'Возвращает сформированную строку в виде prefix & val & postfix. В случае если val=ignore, возвращает ifEmpty
'В случае невозможности преобразования val к строке возвращает ifEmpty
'Самостоятельно образщается к A, где elemID - код данных получаемых анализатором моделей
Dim val As Variant
    
    On Error GoTo ex
    
    val = A.Result(elemID)
    If val = ignore Then
        GetReadyStringA = ifEmpty
    Else
        GetReadyStringA = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyStringA = ifEmpty
End Function

Public Sub fixAllGFSShapesC()
'Замена английской C на русскую С
Dim shp As Visio.Shape
    
    For Each shp In A.gfsShapes
        SetCellVal shp, "Prop.Unit", Replace(cellVal(shp, "Prop.Unit", visUnitsString), "C", "С")
    Next shp
End Sub

Public Function ClearString(ByVal txt As String) As String
'Функция получает на вход строку и возвращает округленное до сотых число (если строка представляла число) или исходную строку
Dim tmpVal As Variant
    On Error Resume Next
    txt = Round(txt, 2)
    ClearString = txt
End Function


'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & "Visio " & Application.Version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.Description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub

