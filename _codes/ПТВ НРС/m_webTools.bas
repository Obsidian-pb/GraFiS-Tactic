Attribute VB_Name = "m_webTools"
Option Explicit
'----------------------Различные инструменты-----------------------


Public Sub GotoWF(ShpObj As Visio.Shape, strMain As String, strAlt As String)
'Переходим на страничку сайта wiki-fire
'strMain - точный адрес тсранички, если указан - переходим по нему
'strAlt - альтернативный адрес странички - используется, если strMain не указан
Const SW_SHOWNORMAL = 1
    
    'По-умолчанию импортируемые из БД значения графис равные NULL заменяются на "0", поэтому корректируем пустые значения на пустую строку ""
    If strMain = "0" Then strMain = ""
    
    'Заменяем косые слэши на нижние подчеркивания, для корректности ссылки
    strAlt = Replace(strAlt, "/", "_")
    
    'Заменяем d strMain пробелы на %20
    strMain = Replace(strMain, " ", "%20")
    strAlt = Replace(strAlt, " ", "%20")
    
    If Len(strMain) > 0 Then
        If InStr(1, strMain, "wiki-fire.org") = 0 Then Exit Sub 'Если в строкой ссылке нет указания на wiki-fire.org - прекращаем выход на страничку - можно выходить ТОЛЬКО на странички wiki-fire.org
        Shell "cmd /cstart " & strMain
    Else
        Shell "cmd /cstart http://wiki-fire.org/" & strAlt & ".ashx"
    End If
End Sub


