Attribute VB_Name = "m_WorkWithStyles"
Option Explicit
'----------------------------В модуле хранятся процедуры работы со стилями--------------------------------------
'----------------------------Наборы стилей документа------------------------------------------------------------


Public Sub StyleExport()
'Узловая процедура обновления стилей
'---Обновляем стили "Пожарная техника"
    If StancilExist("Пожарная техника.vss") Then Refresh_PT
'---Обновляем стили "Техника прочее"
    If StancilExist("Техника прочее.vss") Then Refresh_PTO
'---Обновляем стили "ПТВ"
    If StancilExist("ПТВ.vss") Then Refresh_PTV
'---Обновляем стили "ГДЗС"
    If StancilExist("ГДЗС.vss") Then Refresh_GDZ
'---Обновляем стили "Линии"
    If StancilExist("Линии.vss") Then Refresh_Line
'---Обновляем стили "Связь и освещение"
    If StancilExist("Связь и освещение.vss") Then Refresh_RL
'---Обновляем стили "Водоснабжение"
    If StancilExist("Водоснабжение.vss") Then Refresh_WaterSource
'---Обновляем стили "Очаг"
    If StancilExist("Очаг.vss") Then Refresh_Fire
'---Обновляем стили "Управление СиС"
    If StancilExist("Управление СиС.vss") Then Refresh_Mngmnt
'---Обновляем стили "Прочее"
    If StancilExist("Прочее.vss") Then Refresh_Other


End Sub


Private Sub Refresh_PT()
'Главная процедура экспорта стилей в трафарет "Пожарная техника"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PT_StyleSet(9) As String 'Пожарная техника

    PT_StyleSet(0) = "Т_ОснОбщ_Контур"
    PT_StyleSet(1) = "Т_АЦ_цистерна"
    PT_StyleSet(2) = "Т_ОснЦел_Буквы"
    PT_StyleSet(3) = "Т_ОснЦел_Контур"
    PT_StyleSet(4) = "Т_ПА_Содержимое"
    PT_StyleSet(5) = "Т_Патрубки"
    PT_StyleSet(6) = "Т_Позывные"
'    PT_StyleSet(7) = "Т_Присп_Контур"
'    PT_StyleSet(8) = "Т_Присп_Полоса"
'    PT_StyleSet(9) = "Т_ПрочаяТехника"
    PT_StyleSet(7) = "Т_Спец_Контур"
    PT_StyleSet(8) = "Т_Количество"


    '---Открываем для записи трафарет
        DocOpenClose "Пожарная техника.vss", 1
    
    '---Обновляем стили трафарета "Пожарная техника"
        Set vO_Stenc = Application.Documents("Пожарная техника.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(PT_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(PT_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Пожарная техника.vss", PT_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Пожарная техника.vss", PT_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Пожарная техника.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_PTO()
'Главная процедура экспорта стилей в трафарет "Техника прочее"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PTO_StyleSet(10) As String 'Техника прочее
    
    PTO_StyleSet(0) = "Т_Мотопомпы"
    PTO_StyleSet(1) = "Т_ОснЦел_Буквы"
    PTO_StyleSet(2) = "Т_Патрубки"
    PTO_StyleSet(3) = "Т_Позывные"
    PTO_StyleSet(4) = "Т_Позывные"
    PTO_StyleSet(5) = "Т_Присп_Контур"
    PTO_StyleSet(6) = "Т_Присп_Полоса"
    PTO_StyleSet(7) = "Т_ПрочаяТехника"
    PTO_StyleSet(8) = "Т_Прочая_Гусеничная"
    PTO_StyleSet(9) = "Т_Прочая_Заполненная"


    '---Открываем для записи трафарет
        DocOpenClose "Техника прочее.vss", 1
        
    '---Обновляем стили трафарета "Пожарная техника"
        Set vO_Stenc = Application.Documents("Техника прочее.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(PTO_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(PTO_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Техника прочее.vss", PTO_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Техника прочее.vss", PTO_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Техника прочее.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_PTV()
'Главная процедура экспорта стилей в трафарет "ПТВ"
Dim vs_StencName As String
'Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim PTV_StyleSet(10) As String 'ПТВ

    PTV_StyleSet(0) = "ПТВ_ГидрОбор"
    PTV_StyleSet(1) = "ПТВ_ГидрПодпись"
    PTV_StyleSet(2) = "ПТВ_Дымососы"
    PTV_StyleSet(3) = "ПТВ_Дымососы_Подпись"
    PTV_StyleSet(4) = "ПТВ_Лестницы"
    PTV_StyleSet(5) = "Т_Позывные"
    PTV_StyleSet(6) = "Р_Вс"
    PTV_StyleSet(7) = "ПТВ_ОТ_Контур"
    PTV_StyleSet(8) = "ПТВ_ОТ_Знак"
    PTV_StyleSet(9) = "ПТВ_Ведро"


    '---Открываем для записи трафарет
        DocOpenClose "ПТВ.vss", 1
    
    '---Обновляем стили трафарета "ПТВ"
        Set vO_Stenc = Application.Documents("ПТВ.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(PTV_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(PTV_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("ПТВ.vss", PTV_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "ПТВ.vss", PTV_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "ПТВ.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_GDZ()
'Главная процедура экспорта стилей в трафарет "ГДЗС"
Dim vs_StencName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim GDZ_StyleSet(5) As String 'ГДЗС

    GDZ_StyleSet(0) = "ПТВ_Дымососы"
    GDZ_StyleSet(1) = "ПТВ_Дымососы_Подпись"
    GDZ_StyleSet(2) = "ГДЗ_Звено"
    GDZ_StyleSet(3) = "ГДЗ_Пост"
    GDZ_StyleSet(4) = "Т_Позывные"


    '---Открываем для записи трафарет
        DocOpenClose "ГДЗС.vss", 1
        
    '---Обновляем стили трафарета "ГДЗС"
        Set vO_Stenc = Application.Documents("ГДЗС.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(GDZ_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(GDZ_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("ГДЗС.vss", GDZ_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "ГДЗС.vss", GDZ_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "ГДЗС.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_Line()
'Главная процедура экспорта стилей в трафарет "Линии"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Line_StyleSet(7) As String 'Линии

    Line_StyleSet(0) = "Р_Вс"
    Line_StyleSet(1) = "Р_Нап"
    Line_StyleSet(2) = "Р_НВ"
    Line_StyleSet(3) = "Р_Подпись"
    Line_StyleSet(4) = "Р_Свищ"
    Line_StyleSet(5) = "Р_Мостик"
    Line_StyleSet(6) = "Т_Позывные"

    '---Открываем для записи трафарет
        DocOpenClose "Линии.vss", 1
        
    '---Обновляем стили трафарета "Линии"
        Set vO_Stenc = Application.Documents("Линии.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(Line_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
            Set vO_Stl = ThisDocument.Styles(Line_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Линии.vss", Line_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Линии.vss", Line_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Линии.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_RL()
'Главная процедура экспорта стилей в трафарет "Связь и освещение"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim RL_StyleSet(6) As String 'Связь и освещение

    RL_StyleSet(0) = "С_Подпись"
    RL_StyleSet(1) = "С_Прож"
    RL_StyleSet(2) = "С_РСт"
    RL_StyleSet(3) = "С_Телф"
    RL_StyleSet(4) = "С_ТлфПодп"
    RL_StyleSet(5) = "Т_Позывные"

    '---Открываем для записи трафарет
        DocOpenClose "Связь и освещение.vss", 1
    
    '---Обновляем стили трафарета "Связь и освещение"
        Set vO_Stenc = Application.Documents("Связь и освещение.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(RL_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(RL_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Связь и освещение.vss", RL_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Связь и освещение.vss", RL_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Связь и освещение.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub


Private Sub Refresh_WaterSource()
'Главная процедура экспорта стилей в трафарет "Водоснабжение"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim WaterSource_StyleSet(7) As String 'Водоснабжение

    WaterSource_StyleSet(0) = "ВИ_Водоем"
    WaterSource_StyleSet(1) = "ВИ_Озеро"
    WaterSource_StyleSet(2) = "ВИ_ОзОбъем"
    WaterSource_StyleSet(3) = "ВИ_ОзПодп1"
    WaterSource_StyleSet(4) = "ВИ_ПГ"
    WaterSource_StyleSet(5) = "ВИ_Подписи"
    WaterSource_StyleSet(6) = "ВИ_Емкость"


    '---Открываем для записи трафарет
        DocOpenClose "Водоснабжение.vss", 1
    
    '---Обновляем стили трафарета "Водоснабжение"
        Set vO_Stenc = Application.Documents("Водоснабжение.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(WaterSource_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(WaterSource_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Водоснабжение.vss", WaterSource_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Водоснабжение.vss", WaterSource_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Водоснабжение.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Fire()
'Главная процедура экспорта стилей в трафарет "Очаг"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Fire_StyleSet(10) As String 'Очаг

    Fire_StyleSet(0) = "О_Задымление"
    Fire_StyleSet(1) = "О_Обрушение"
    Fire_StyleSet(2) = "О_ОдПожар"
    Fire_StyleSet(3) = "О_ОдПожПодп"
    Fire_StyleSet(4) = "О_Очаг"
    Fire_StyleSet(5) = "О_ПлощПож"
    Fire_StyleSet(6) = "О_ПлощПожПодп"
    Fire_StyleSet(7) = "О_ПлощТушПодп"
    Fire_StyleSet(8) = "О_Постр"
    Fire_StyleSet(9) = "О_Распр"

    '---Открываем для записи трафарет
        DocOpenClose "Очаг.vss", 1
    
    '---Обновляем стили трафарета "Очаг"
        Set vO_Stenc = Application.Documents("Очаг.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(Fire_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(Fire_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Очаг.vss", Fire_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Очаг.vss", Fire_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Очаг.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Mngmnt()
'Главная процедура экспорта стилей в трафарет "Управление СиС"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Mngmnt_StyleSet(3) As String 'Управление СиС

    Mngmnt_StyleSet(0) = "О_РНД"
    Mngmnt_StyleSet(1) = "УС_УТП"
    Mngmnt_StyleSet(2) = "УС_Штаб"

    '---Открываем для записи трафарет
        DocOpenClose "Управление СиС.vss", 1
    
    '---Обновляем стили трафарета "Управление СиС"
        Set vO_Stenc = Application.Documents("Управление СиС.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(Mngmnt_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(Mngmnt_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Управление СиС.vss", Mngmnt_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Управление СиС.vss", Mngmnt_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Управление СиС.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

Private Sub Refresh_Other()
'Главная процедура экспорта стилей в трафарет "Прочее"
Dim vs_StencName As String
Dim vs_StyleName As String
Dim vO_Stenc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
Dim Other_StyleSet(1) As String 'Прочее

    Other_StyleSet(0) = "Пр_Подписи"
'    Mngmnt_StyleSet(1) = "УС_УТП"
'    Mngmnt_StyleSet(2) = "УС_Штаб"


    '---Открываем для записи трафарет
        DocOpenClose "Прочее.vss", 1
    
    '---Обновляем стили трафарета "Прочее"
        Set vO_Stenc = Application.Documents("Прочее.vss")
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(Other_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(Other_StyleSet(i))
        '---Проверяем есть ли указаный стиль vs_StyleName в трафарете vs_StencName
            If StyleExist("Прочее.vss", Other_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh "Прочее.vss", Other_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиль vs_StyleName в трафарет vs_StencName
                vO_Stenc.Drop vO_Stl, 0, 0
            End If
        Next i
    '---Закрываем для записи трафарет
        DocOpenClose "Прочее.vss", 0

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing
End Sub

'-------------------------------------Процедуры обновления-------------------------------------------------------
Private Sub StyleRefresh(as_StnclName As String, as_StyleName As String)
'Процедура обновления стиля as_StyleName в трафарете as_StnclName
Dim vO_Stenc As Visio.Document
Dim vO_StyleFrom As Visio.style
Dim vO_StyleTo As Visio.style
Dim vs_RowName As String

'---Создвем необходимый набор объектов
    Set vO_Stenc = Application.Documents(as_StnclName)
    Set vO_StyleFrom = ThisDocument.Styles(as_StyleName)
    Set vO_StyleTo = vO_Stenc.Styles(as_StyleName)

'---Обновляем секции для стиля "Текст"
    If vO_StyleFrom.Cells("EnableTextProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("Char.Font").FormulaU = vO_StyleFrom.Cells("Char.Font").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Size").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Size").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.FontScale").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.FontScale").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Letterspace").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Letterspace").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Color").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.Color").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.ColorTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("Char.ColorTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Char.Style").FormulaU = vO_StyleFrom.Cells("Char.Style").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Case").FormulaU = vO_StyleFrom.Cells("Char.Case").ResultStr(visUnitsString)
        vO_StyleTo.Cells("Char.Pos").FormulaU = vO_StyleFrom.Cells("Char.Pos").ResultStr(visUnitsString)
    End If
    
'---Обновляем секции для стиля "Линия"
    If vO_StyleFrom.Cells("EnableLineProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("LinePattern").FormulaU = vO_StyleFrom.Cells("LinePattern").ResultStr(visUnitsString)
'        vO_StyleTo.Cells("LineWeight").FormulaU = vO_StyleFrom.Cells("LineWeight").ResultStr(visUnitsString)
        vO_StyleTo.Cells("LineColor").FormulaU = Chr(34) & vO_StyleFrom.Cells("LineColor").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("LineCap").FormulaU = vO_StyleFrom.Cells("LineCap").ResultStr(visUnitsString)
        vO_StyleTo.Cells("LineColorTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("LineColorTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("Rounding").FormulaU = Chr(34) & vO_StyleFrom.Cells("Rounding").ResultStr(visUnitsString) & Chr(34)
    End If
    
'---Обновляем секции для стиля "Заливка"
    If vO_StyleFrom.Cells("EnableFillProps").Result(visNumber) = 1 Then
        vO_StyleTo.Cells("FillForegnd").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillForegnd").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillForegndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillForegndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillBkgnd").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillBkgnd").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillBkgndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("FillBkgndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("FillPattern").FormulaU = vO_StyleFrom.Cells("FillPattern").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwForegnd").FormulaU = vO_StyleFrom.Cells("ShdwForegnd").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwForegndTRans").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShdwForegndTRans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShdwBkgnd").FormulaU = vO_StyleFrom.Cells("ShdwBkgnd").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShdwBkgndTrans").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShdwBkgndTrans").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShdwPattern").FormulaU = vO_StyleFrom.Cells("ShdwPattern").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShapeShdwOffsetX").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwOffsetX").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwOffsetY").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwOffsetY").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwType").FormulaU = vO_StyleFrom.Cells("ShapeShdwType").ResultStr(visUnitsString)
        vO_StyleTo.Cells("ShapeShdwObliqueAngle").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwObliqueAngle").ResultStr(visUnitsString) & Chr(34)
        vO_StyleTo.Cells("ShapeShdwScaleFactor").FormulaU = Chr(34) & vO_StyleFrom.Cells("ShapeShdwScaleFactor").ResultStr(visUnitsString) & Chr(34)
    End If
    

'---Очищаем объекты
    Set vO_StyleFrom = Nothing
    Set vO_StyleTo = Nothing
    Set vO_Stenc = Nothing
End Sub


Private Sub DocOpenClose(asS_StencName As String, asb_OpenClose As Byte)
'Процедура открытия-закрытия документа для изменений.  0 - закрыть; 1 - открыть
Dim doc As Visio.Document
Dim pth As String

'---Идентифицируем необходимый документ
Set doc = Documents(asS_StencName)
pth = doc.fullName

'---Открываем или закрываем документ в зависимости от указанного asb_OpenClose
    If asb_OpenClose = 0 Then
        If doc.ReadOnly = False Then
            doc.Close
            Application.Documents.OpenEx pth, visOpenRO + visOpenDocked
        End If
    Else
        If doc.ReadOnly = True Then
            doc.Close
            Application.Documents.OpenEx pth, visOpenRW + visOpenDocked
        End If
    End If

'---Очищаем объекты
Set doc = Nothing

End Sub


'-----------------------------------------------------Функции--------------------------------------------------
Private Function StyleExist(asS_StencName As String, as_StyleName As String) As Boolean
'Функция возвращает булево значение наличия данного стиля в указанном документе
Dim vO_Style As Visio.style

StyleExist = False

For Each vO_Style In Application.Documents(asS_StencName).Styles
    If vO_Style.Name = as_StyleName Then StyleExist = True
Next vO_Style

End Function

Private Function StancilExist(asS_StencName As String) As Boolean
'Функция возвращает булево значение наличия данного трафарета
Dim vO_Stencil As Visio.Document
StancilExist = False

For Each vO_Stencil In Application.Documents
    If vO_Stencil.Name = asS_StencName Then StancilExist = True
Next vO_Stencil

End Function


