VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ButEventOpenWater As ClassLake
Dim WithEvents WaterAppEvents As Visio.Application
Attribute WaterAppEvents.VB_VarHelpID = -1


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура закрытия документа и удаления его рабочих элементов

'---Очищаем объекты ButEvent и WaterAppEvents и удаляем кнопку "Обратить в естественный водоисточник" с панели управления "Превращения"
    Set ButEventOpenWater = Nothing
    Set WaterAppEvents = Nothing
    DeleteButtons
    
'---В случае, если на панели "Превращения нет ни одной кнопки, удаляем её
    If Application.CommandBars("Превращения").Controls.Count = 0 Then RemoveTBImagination

End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo Tail
    
'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells
    
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True

'---Импортируем мастера
    MastersImport
    
'---Добавляем ячейку Аспект (если еще не была добавлена)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---Создаем панель управления "Превращения" и добавляем на нее кнопку "Обратить в откртый водоисточник"
    AddTBImagination
    AddButtons

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Определяем управляющие объекты
    Set WaterAppEvents = Visio.Application
    Set ButEventOpenWater = New ClassLake
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "Программа вызвала ошибку! Если это будет повторяться, свяжитесь с разработчиком."
End Sub


Private Sub WaterAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer
'---Проверяем имя ячейки
'    ShpInd = Cell.Shape.ID
    
    If cell.Name = "Prop.PipeType" Or cell.Name = "Prop.PipeDiameter" Or cell.Name = "Prop.Pressure" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКА диаметров
'        DiametersListImport (ShpInd)
        '---Запускаем процедуру получения СПИСКА напоров
        PressuresListImport (ShpInd)
        '---Запускаем процедуру пересчета водоотдачи
        ProductionImport (ShpInd)
    End If

'В случае, если произошло изменение не нужной ячейки прекращаем событие
End Sub

Public Sub MastersImport()
'---Импортируем мастера
    MasterImportSub "Водоснабжение.vss", "Водоем1_Мелкий"
    MasterImportSub "Водоснабжение.vss", "Водоем1_Средний"
'    MasterImportSub "Водоснабжение.vss", "Водоем1_Средний" ВИ_Емкость
    
End Sub

Private Sub WaterAppEvents_ShapeAdded(ByVal Shape As IVShape)
'Событие добавления на лист фигуры
Dim v_Cntrl As CommandBarControl

'---Включаем обработку ошибок
    On Error GoTo Tail

'---Проверяем включена ли кнопка водоема и в зависмости от этого пытаемся обращатить вброшенную фигуру
    Set v_Ctrl = Application.CommandBars("Превращения").Controls("Естественный водоисточник")
        If v_Ctrl.State = msoButtonDown Then
'            If v_Ctrl.State = msoButtonUp Then
                If IsSelectedOneShape(False) Then
                '---Если выбрана хоть одна фигура - пытаемся ее обратить
                    If IsHavingUserSection(False) And IsSquare(False) Then
                    '---Обращаем фигуру в фигуру зону горения
                        ButEventOpenWater.MorphToLake
                    End If
'                Else
'                    PS_CheckButtons v_Ctrl
                End If
'            End If
        
        
        
            '---Запускаем процедуру обращения в естественный водоем
'            ButEventOpenWater.MorphToLake
        End If

Set v_Ctrl = Nothing
Exit Sub
Tail:
    SaveLog Err, "WaterAppEvents_ShapeAdded"
    MsgBox "Программа вызвала ошибку! Если это будет повторяться, свяжитесь с разработчиком."
    Set v_Ctrl = Nothing
End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If

End Sub





