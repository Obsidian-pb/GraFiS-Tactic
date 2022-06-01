VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents PTVAppEvents As Visio.Application
Attribute PTVAppEvents.VB_VarHelpID = -1

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура при закрытии документа

'---Очищаем переменную приложения
Set PTVAppEvents = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'Процедура при открытии документа
    
    On Error GoTo EX
    
'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set PTVAppEvents = Visio.Application
    
''---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates
    
Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub PTVAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer
'---Проверяем имя ячейки
'MsgBox Cell.Name
    
    On Error GoTo Tail
    
    If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
            Or cell.Name = "Prop.StreamType" Or cell.Name = "Prop.Head" Then
        ShpInd = cell.Shape.ID
        'Проверяем, необходимо ли обновлять значения из БД
        If cell.Shape.Cells("Prop.TTHType").ResultStr(visString) = "По модели ствола" Then
            '---Запускаем процедуру получения СПИСКОВ стволов
            If cell.Name = "Prop.TTHType" Then
                StvolModelsListImport (ShpInd)
            End If
            
            '---Запускаем процедуру получения списков ВАРИАНТОВ стволов
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Then
                StvolVariantsListImport (ShpInd)
            End If
            
            '---Запускаем процедуру получения Кратности для пенных стволов в соответствии с его моделью
            StvolRFImport (ShpInd)
            
            '---Запускаем процедуру получения списков ВИДОВ СТРУЙ стволов
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" Then
                StvolStreamTypesListImport (ShpInd)
            End If
            
            '---Запускаем процедуру получения Условного прохода ствола в соответствии с его ВИДОМ СТРУИ
            StvolDiameterInImport (ShpInd)
            
            '---Запускаем процедуру получения списка НАПОРОВ для данного вида струи стволов
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Then
                StvolHeadListImport (ShpInd)
            End If
            
            '---Запускаем процедуру получения ТИПА СТРУи ствола в соответствии с его ВИДОМ СТРУИ (компактная, распыленная или иная)
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Then
                StvolStreamValueImport (ShpInd)
            End If
            
            '---Запускаем процедуру пересчета расхода воды из ствола
            If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Or cell.Name = "Prop.Variant" _
                Or cell.Name = "Prop.StreamType" Or cell.Name = "Prop.Head" Then
                StvolProductionImport (ShpInd)
            End If
            
            
        End If
    End If
    
    '---Получаем ссылку на страничку в wiki-fire.org (только для стволов!!!)
    If cell.Name = "Prop.TTHType" Or cell.Name = "Prop.StvolType" Then
        If cell.Shape.Cells("Prop.TTHType").ResultStr(visString) = "По модели ствола" Then
            '---Запускаем проку получения ссылки на wiki-fire.org
            StvolWFLinkImport (ShpInd)
            StvolHeadDiapasoneImport (ShpInd)
        Else
            StvolWFLinkFree (ShpInd)
        End If
    End If
    
    If cell.Name = "Prop.WEType" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКОВ гидроэлеваторов
        WEModelsListImport (ShpInd)
        '---Запускаем процедуру пересчета расхода воды из ствола
        StvolProductionImport (ShpInd)
    End If
    
    If cell.Name = "Prop.WFType" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКОВ Всасывающих сеток
        WFModelsListImport (ShpInd)
        '---Запускаем процедуру пересчета производительности всасывающих сеток
        StvolProductionImport (ShpInd)
    End If
    
    If cell.Name = "Prop.ColPressure" Or cell.Name = "Prop.Patr" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру пересчета производительности колонок
        ColFlowMaxImport (ShpInd)
    End If
    
'В случае, если произошло изменение не нужной ячейки прекращаем событие
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "PTVAppEvents_CellChanged"
End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell

    On Error GoTo Tail

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "AddTimeUserCells"
End Sub

'============
'Код для показа мастера:
'ThisDocument.Masters("Гидроабразивный ствол").Hidden = False
