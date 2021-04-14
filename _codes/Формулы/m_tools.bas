Attribute VB_Name = "m_tools"


Public Function ClearString(ByVal txt As String) As String
'Функция получает на вход строку и возвращает округленное до сотых число (если строка представляла число) или исходную строку
Dim tmpVal As Variant
    On Error Resume Next
    txt = Round(txt, 2)
    ClearString = txt
End Function

Public Sub sleep(ByVal sec As Single, Optional ByVal doE As Boolean = False)
Dim i As Long
Dim endTime As Single
    
    endTime = DateTime.Timer + sec
    Do While DateTime.Timer < endTime
        If doE Then DoEvents
    Loop
    
'    Debug.Print "sleep 0.5"
    
End Sub

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

