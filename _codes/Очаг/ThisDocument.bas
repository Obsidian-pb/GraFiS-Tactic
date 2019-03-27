VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private dbs As Database
Dim WithEvents SquareAppEvents As Visio.Application
Attribute SquareAppEvents.VB_VarHelpID = -1

Dim ButEventFireArea As ClassFireArea, ButEventStorm As ClassStorm, ButEventFog As ClassFog, ButEventRush As ClassRush

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo EX
    
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

'---Инициируем объект SquareAppEvents для реагирования на действия пользователя
    Set SquareAppEvents = Visio.Application
    
'---Создаем панель управления "Превращения" и добавляем на нее кнопку "Обратить в зону горения"
    AddTBImagination
    AddButtons

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Включаем показ окон
    VfB_NotShowPropertiesWindow = False

'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

Set ButEventFireArea = New ClassFireArea
Set ButEventStorm = New ClassStorm
Set ButEventFog = New ClassFog
Set ButEventRush = New ClassRush

'---Добавляем свойство документа "FireTime"
    sm_AddFireTime
Exit Sub
EX:
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура закрытия документа и удаления его рабочих элементов

'---Очищаем объект ButEvent и удаляем кнопку "Площадь" с панели управления "Превращения"
    Set ButEventFireArea = Nothing
    Set ButEventStorm = Nothing
    Set ButEventFog = Nothing
    Set ButEventRush = Nothing
    DeleteButtons
    
'---В случае, если на панели "Превращения нет ни одной кнопки, удаляем её
    If Application.CommandBars("Превращения").Controls.Count = 0 Then RemoveTBImagination
'---Очищаем переменную приложения
    Set SquareAppEvents = Nothing
    
End Sub


Private Sub SquareAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Long '(64) - Площадь пожара
'---Проверяем имя ячейки
    
    
    If cell.Name = "Prop.FireCategorie" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКОВ описаний объектов пожара для указанной категории
        DescriptionsListImport (ShpInd)
    End If
        
    If cell.Name = "Prop.FireDescription" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения ЗНАЧЕНИЙ факторов пожара для данного описания
        GetFactorsByDescription (ShpInd)
    End If
        
        
'    ElseIf Cell.Name = "Prop.FireObject" Then
'        '---Запускаем процедуру получения ЗНФЧЕНИЙ интенсивностей подачи воды объектов пожара
'        GetIntenseWaterByObject (ShpInd)
'        '---Запускаем процедуру получения ЗНФЧЕНИЙ линейной скорости для объектов пожара
'        GetSpeedByObject (ShpInd)
'    ElseIf Cell.Name = "Prop.FireMaterials" Then
'        '---Запускаем процедуру получения ЗНФЧЕНИЙ интенсивностей подачи воды горючих материалов
'        GetIntenseWaterByMaterial (ShpInd)
'        '---Запускаем процедуру получения ЗНФЧЕНИЙ линейной скорости для материалов пожара
'        GetSpeedByMaterial (ShpInd)
'    End If
    
    
    
'MsgBox Cell.Shape.Index
'MsgBox Cell.Shape.ID

'В случае, если произошло изменение не нужной ячейки прекращаем событие
End Sub

Public Sub MastersImport()
'---Импортируем мастера
'Dim mstr As Visio.Master

    MasterImportSub "Очаг.vss", "Задымление1_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление2_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление3_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление4_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление5_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление6_Мелкий"
    MasterImportSub "Очаг.vss", "Задымление1_Средний"
    MasterImportSub "Очаг.vss", "Задымление2_Средний"
    MasterImportSub "Очаг.vss", "Задымление3_Средний"
    MasterImportSub "Очаг.vss", "Задымление4_Средний"
    MasterImportSub "Очаг.vss", "Задымление5_Средний"
    MasterImportSub "Очаг.vss", "Задымление6_Средний"
    MasterImportSub "Очаг.vss", "Задымление1_Крупный"
    MasterImportSub "Очаг.vss", "Задымление2_Крупный"
    MasterImportSub "Очаг.vss", "Задымление3_Крупный"
    MasterImportSub "Очаг.vss", "Задымление4_Крупный"
    MasterImportSub "Очаг.vss", "Задымление5_Крупный"
    MasterImportSub "Очаг.vss", "Задымление6_Крупный"
    MasterImportSub "Очаг.vss", "Очаг1_Мелкий"
    MasterImportSub "Очаг.vss", "Очаг2_Мелкий"
    MasterImportSub "Очаг.vss", "Очаг3_Мелкий"
    MasterImportSub "Очаг.vss", "Очаг4_Мелкий"
    MasterImportSub "Очаг.vss", "Очаг1_Средний"
    MasterImportSub "Очаг.vss", "Очаг2_Средний"
    MasterImportSub "Очаг.vss", "Очаг3_Средний"
    MasterImportSub "Очаг.vss", "Очаг4_Средний"
    MasterImportSub "Очаг.vss", "Очаг1_Крупный"
    MasterImportSub "Очаг.vss", "Очаг2_Крупный"
    MasterImportSub "Очаг.vss", "Очаг3_Крупный"
    MasterImportSub "Очаг.vss", "Очаг4_Крупный"
    MasterImportSub "Очаг.vss", "Огненный шторм"
    MasterImportSub "Очаг.vss", "Обрушение"

End Sub



Private Sub sm_AddFireTime()
'Процедура добавляет в документ свойство - время начала пожара, в случае его отсутствия

    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If

End Sub

Private Sub SquareAppEvents_ShapeAdded(ByVal Shape As IVShape)
'Событие добавления на лист фигуры
Dim v_Cntrl As CommandBarControl
'Dim SecExists As Boolean
    
'---Включаем обработку ошибок
    On Error GoTo Tail
    
'---Проверяем является ли добавленная фигура незамкнутой линией без свойств
'    SecExists = Shape.SectionExists(visSectionProp, 0)
'    If Shape.AreaIU > 0 Or SecExists Then Exit Sub

'---Проверяем какая кнопка нажата и в зависимости от этого выполняем действие
    For Each v_Ctrl In Application.CommandBars("Превращения").Controls
        If v_Ctrl.State = msoButtonDown Then
            Select Case v_Ctrl.Caption
                Case Is = "Площадь"
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---Запускаем процедуру обращения в зону горения
                        ButEventFireArea.MorphToFireArea
                        End If
                    End If
                Case Is = "Задымление"
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---Запускаем процедуру обращения в задымление
                        ButEventFog.MorphToFog
                        End If
                    End If
                Case Is = "Обрушение"
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---Запускаем процедуру обращения в обрушение
                        ButEventRush.MorphToRush
                        End If
                    End If
                Case Is = "Шторм"
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If IsHavingUserSection(False) And IsSquare(False) Then
                        '---Запускаем процедуру обращения в шторм
                        ButEventStorm.MorphToStorm
                        End If
                    End If
            End Select
        End If
    Next v_Ctrl
    
Exit Sub
Tail:
'    MsgBox Err.Description
    MsgBox "В ходе работы программы возникла ошибка! Если она будет повторяться - обратитесь к разработчику."
    SaveLog Err, "Document_DocumentOpened"
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


