Attribute VB_Name = "Tools"




Public Function CallToStr(ByVal a_Call As String) As String
'Функция получает строку с позывным, и если ее можно привести к цифре - приводит, возвращает строковое значение
On Error Resume Next
    a_Call = Int(a_Call)
    CallToStr = a_Call
End Function

