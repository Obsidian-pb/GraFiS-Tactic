Attribute VB_Name = "Vichislenia"
Option Explicit

Sub SquareSet(ShpObj As Visio.Shape)
'Процедура присвоения текстовому полю выделенной фигуры значения площади фигуры
'Только для фигур Площади пожара
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 'переводим из квадратных дюймов в квадратные метры
ShpObj.Cells("User.FireSquareP").FormulaForceU = SquareCalc

End Sub

Sub USquareSet(ShpObj As Visio.Shape)
'Процедура присвоения текстовому полю выделенной фигуры значения площади фигуры
Dim SquareCalc As Long

SquareCalc = ShpObj.AreaIU * 0.00064516 'переводим из квадратных дюймов в квадратные метры
ShpObj.Cells("User.SquareP").FormulaForceU = SquareCalc

End Sub


Sub s_SetFireTime(ShpObj As Visio.Shape, Optional showDoCmd As Boolean = True)
'Процедура присвоения ячейке документа User.FireTime значения времени указанного при вбрасывании фигуры "Очаг"
Dim vD_CurDateTime As Double

On Error Resume Next

'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Присваиваем значению времени возникновения пожара текущее значение
            vD_CurDateTime = Now()
            ShpObj.Cells("Prop.FireTime").FormulaU = _
                "DATETIME(" & str(vD_CurDateTime) & ")"
        
        '---Показываем окно свойств фигуры
            If showDoCmd Then Application.DoCmd (1312)
            
        '---Если в Шэйп-листе документа отсутствует строка "User.FireTime", создаем её
            If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
                Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
            End If
            
        '---Переносим новые данные из шейп личста фигуры в шейп лист документа
            Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
                "DATETIME(" & str(CDbl(ShpObj.Cells("Prop.FireTime").Result(visDate))) & ")"
    Else
        '---Показываем окно свойств фигуры
            If showDoCmd Then Application.DoCmd (1312)
    End If
    '---Добавляем в перечень свойств страницы данные о временах реагирования
    AddPageTimeProps ShpObj

End Sub

Public Sub ShowTimesForm(ByRef shp As Visio.Shape)
    F_Times.ShowMe shp
End Sub













'------------------Создание строк данных в странице----------------------
Public Sub AddPageTimeProps(ByRef shpFire As Visio.Shape)
Dim tmpRowInd As Integer
'Dim tmpRow As Visio.Row
Dim shp As Visio.Shape

    
'    On Error Resume Next
    
    Set shp = Application.ActivePage.PageSheet
    
    'Время возникновения
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время возникновения" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время возникновения пожара" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = 5
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 20
        SetCellVal shp, "Prop.FireTime", CellVal(shpFire, "Prop.FireTime", visDate) ' GetVal(fir)

        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время возникновения" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireTime" & Chr(34) & ",TheDoc!User.FireTime)+DEPENDSON(TheDoc!User.FireTime)"  'ВАЖНО !!! формулы вставлялись явным образом - текстом, иначе можно напороться на ошибки несовпадения индексов строк Scratch
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireTime" & Chr(34) & ", Prop.FireTime) + DEPENDSON(Prop.FireTime)"      'ВАЖНО !!! формулы вставлялись явным образом - текстом, иначе можно напороться на ошибки несовпадения индексов строк Scratch

    'Время обнаружения
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FindTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время обнаружения" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время обнаружения пожара" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FindTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 21
'        SetCellVal shp, "Prop.FindTime", GetVal(fin)
        SetCellVal shp, "Prop.FindTime", CellVal(shpFire, "Prop.FindTime", visDate)

        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время обнаружения" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FindTime" & Chr(34) & ",TheDoc!User.FindTime)+DEPENDSON(TheDoc!User.FindTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FindTime" & Chr(34) & ", Prop.FindTime) + DEPENDSON(Prop.FindTime)"
        
    'Время сообщения
        tmpRowInd = shp.AddNamedRow(visSectionProp, "InfoTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время сообщения" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время сообщения о пожаре" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.InfoTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 22
'        SetCellVal shp, "Prop.InfoTime", GetVal(inf)
        SetCellVal shp, "Prop.InfoTime", CellVal(shpFire, "Prop.InfoTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время сообщения" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.InfoTime" & Chr(34) & ",TheDoc!User.InfoTime)+DEPENDSON(TheDoc!User.InfoTime)"  '
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.InfoTime" & Chr(34) & ", Prop.InfoTime) + DEPENDSON(Prop.InfoTime)"

    'Время прибытия первого подразделения
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstArrivalTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время прибытия" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время прибытия к месту пожара первого подразделения" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FirstArrivalTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 23
'        SetCellVal shp, "Prop.FirstArrivalTime", GetVal(fArr)
        SetCellVal shp, "Prop.FirstArrivalTime", CellVal(shpFire, "Prop.FirstArrivalTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время прибытия" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstArrivalTime" & Chr(34) & ",TheDoc!User.FirstArrivalTime)+DEPENDSON(TheDoc!User.FirstArrivalTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstArrivalTime" & Chr(34) & ", Prop.FirstArrivalTime) + DEPENDSON(Prop.FirstArrivalTime)"

    'Время подачи первого ствола
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FirstStvolTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время подачи ствола" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время подачи первого ствола" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FirstStvolTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 24
'        SetCellVal shp, "Prop.FirstStvolTime", GetVal(fArr)
        SetCellVal shp, "Prop.FirstStvolTime", CellVal(shpFire, "Prop.FirstStvolTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время подачи ствола" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FirstStvolTime" & Chr(34) & ",TheDoc!User.FirstStvolTime)+DEPENDSON(TheDoc!User.FirstStvolTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FirstStvolTime" & Chr(34) & ", Prop.FirstStvolTime) + DEPENDSON(Prop.FirstStvolTime)"

    'Время локализации
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LocalizationTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время локализации" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время локализации" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LocalizationTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 25
'        SetCellVal shp, "Prop.LocalizationTime", GetVal(fArr)
        SetCellVal shp, "Prop.LocalizationTime", CellVal(shpFire, "Prop.LocalizationTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время локализации" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LocalizationTime" & Chr(34) & ",TheDoc!User.LocalizationTime)+DEPENDSON(TheDoc!User.LocalizationTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LocalizationTime" & Chr(34) & ", Prop.LocalizationTime) + DEPENDSON(Prop.LocalizationTime)"
        
    'Время ЛОГ
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LOGTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время ликвидации ОГ" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации открытого горения" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LOGTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 26
'        SetCellVal shp, "Prop.LOGTime", GetVal(fArr)
        SetCellVal shp, "Prop.LOGTime", CellVal(shpFire, "Prop.LOGTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время ликвидации ОГ" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LOGTime" & Chr(34) & ",TheDoc!User.LOGTime)+DEPENDSON(TheDoc!User.LOGTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LOGTime" & Chr(34) & ", Prop.LOGTime) + DEPENDSON(Prop.LOGTime)"
        
    'Время ЛПП
        tmpRowInd = shp.AddNamedRow(visSectionProp, "LPPTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время ликвидации ПП" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации последствий пожара" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.LPPTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 27
'        SetCellVal shp, "Prop.LPPTime", GetVal(fArr)
        SetCellVal shp, "Prop.LPPTime", CellVal(shpFire, "Prop.LPPTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время ликвидации ПП" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.LPPTime" & Chr(34) & ",TheDoc!User.LPPTime)+DEPENDSON(TheDoc!User.LPPTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.LPPTime" & Chr(34) & ", Prop.LPPTime) + DEPENDSON(Prop.LPPTime)"
        
    'Время окончания пожара
        tmpRowInd = shp.AddNamedRow(visSectionProp, "FireEndTime", visTagDefault)
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsLabel).Formula = """" & "Время завершения работ" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsPrompt).Formula = """" & "Полное астрономическое время ликвидации последствий пожара" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsType).FormulaU = "IF(STRSAME(Prop.FireEndTime," & Chr(34) & Chr(34) & "),0,5)"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsFormat).FormulaU = """" & "{{dd.MM.yyyy H:mm}}" & """"
        shp.CellsSRC(visSectionProp, tmpRowInd, visCustPropsSortKey).FormulaU = 28
'        SetCellVal shp, "Prop.FireEndTime", GetVal(fArr)
        SetCellVal shp, "Prop.FireEndTime", CellVal(shpFire, "Prop.FireEndTime", visDate)
        
        tmpRowInd = shp.AddRow(visSectionScratch, visRowLast, visTagDefault)
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchA).FormulaU = """" & "Время завершения работ" & """"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchC).FormulaU = "SETF(" & Chr(34) & "Prop.FireEndTime" & Chr(34) & ",TheDoc!User.FireEndTime)+DEPENDSON(TheDoc!User.FireEndTime)"
        shp.CellsSRC(visSectionScratch, tmpRowInd, visScratchD).FormulaU = "SETF(" & Chr(34) & "TheDoc!User.FireEndTime" & Chr(34) & ", Prop.FireEndTime) + DEPENDSON(Prop.FireEndTime)"
        
End Sub


