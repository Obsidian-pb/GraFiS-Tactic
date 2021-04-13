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
