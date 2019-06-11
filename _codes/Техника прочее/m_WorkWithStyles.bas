Attribute VB_Name = "m_WorkWithStyles"
'----------------------------В модуле хранятся процедуры работы со стилями--------------------------------------
Dim PTV_StyleSet(9) As String


Private Sub s_StyleSetsDeclare()
'Процедура задает перечень стилей для обновления
'---Стили "Пожарная техника"
    PTV_StyleSet(0) = "Т_Мотопомпы"
    PTV_StyleSet(1) = "Т_ОснЦел_Буквы"
    PTV_StyleSet(2) = "Т_Патрубки"
    PTV_StyleSet(3) = "Т_Позывные"
    PTV_StyleSet(4) = "Т_Присп_Контур"
    PTV_StyleSet(5) = "Т_Присп_Полоса"
    PTV_StyleSet(6) = "Т_Прочая_Гусеничная"
    PTV_StyleSet(7) = "Т_Прочая_Заполненная"
    PTV_StyleSet(8) = "Т_ПрочаяТехника"

    
    
End Sub


Public Sub StyleExport()
'Главная процедура экспорта стилей в активный документ
Dim vO_Doc As Visio.Document
Dim vO_Stl As Visio.style
Dim i As Integer
    
'---Создаем набор названий стилей
    s_StyleSetsDeclare

'---Обновляем стили активного документа
    Set vO_Doc = Application.ActiveDocument
    
    '---Перебираем все стили трафарета
        For i = 0 To UBound(PTV_StyleSet()) - 1
        '---Выбираем очередной стиль трафарета
        Set vO_Stl = ThisDocument.Styles(PTV_StyleSet(i))
        '---Проверяем есть ли указаный стиль PT_StyleSet(i) в активном документе
            If StyleExist(PTV_StyleSet(i)) Then
            '---Если есть - обновляем его
                StyleRefresh PTV_StyleSet(i)
            Else
            '---Если нет - вбрасываем стиь PT_StyleSet(i) в активный документе
                vO_Doc.Drop vO_Stl, 0, 0
            End If
        Next i

'---Очищаем объекты
    Set vO_Stenc = Nothing
    Set vO_Stl = Nothing

End Sub


Private Sub StyleRefresh(as_StyleName As String)
'Процедура обновления стиля as_StyleName в активном документе
Dim vO_StyleFrom As Visio.style
Dim vO_StyleTo As Visio.style
Dim vs_RowName As String

    On Error GoTo EX

'---Создвем необходимый набор объектов
    Set vO_StyleFrom = ThisDocument.Styles(as_StyleName)
    Set vO_StyleTo = ActiveDocument.Styles(as_StyleName)

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
Exit Sub
EX:
'---Очищаем объекты
    Set vO_StyleFrom = Nothing
    Set vO_StyleTo = Nothing
    Set vO_Stenc = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "StyleRefresh"
End Sub


Private Function StyleExist(as_StyleName As String) As Boolean
'Функция возвращает булево значение наличия данного стиля в активном документе
Dim vO_Style As Visio.style

    StyleExist = False
    
    For Each vO_Style In Application.ActiveDocument.Styles
        If vO_Style.Name = as_StyleName Then StyleExist = True
    Next vO_Style

End Function



