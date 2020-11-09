VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 1000

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
    
'---Активируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If

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
    
'---Деактивируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then Set app = Nothing
    
'---В случае, если на панели "Превращения нет ни одной кнопки, удаляем её
    If Application.CommandBars("Превращения").Controls.Count = 0 Then RemoveTBImagination
'---Очищаем переменную приложения
    Set SquareAppEvents = Nothing
    
End Sub


Private Sub SquareAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Long '(64) - Площадь пожара
'Dim shpFire As Visio.Shape 'Фигура Площадь пожара

'---Проверяем не произошло ли событие в мастере
    On Error GoTo EX
    If cell.Shape.ContainingMasterID >= 0 Then Exit Sub
    
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
    
    If cell.Name = "Prop.FireTime" Then
        '---Переносим новые данные из шейп личста фигуры в шейп лист документа
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = _
            "DATETIME(" & str(CDbl(cell.Shape.Cells("Prop.FireTime").Result(visDate))) & ")"
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
EX:
End Sub

Public Sub MastersImport()
'---Импортируем мастера
'Dim mstr As Visio.Master

    MasterImportSub "Задымление1_Мелкий"
    MasterImportSub "Задымление2_Мелкий"
    MasterImportSub "Задымление3_Мелкий"
    MasterImportSub "Задымление4_Мелкий"
    MasterImportSub "Задымление5_Мелкий"
    MasterImportSub "Задымление6_Мелкий"
    MasterImportSub "Задымление1_Средний"
    MasterImportSub "Задымление2_Средний"
    MasterImportSub "Задымление3_Средний"
    MasterImportSub "Задымление4_Средний"
    MasterImportSub "Задымление5_Средний"
    MasterImportSub "Задымление6_Средний"
    MasterImportSub "Задымление1_Крупный"
    MasterImportSub "Задымление2_Крупный"
    MasterImportSub "Задымление3_Крупный"
    MasterImportSub "Задымление4_Крупный"
    MasterImportSub "Задымление5_Крупный"
    MasterImportSub "Задымление6_Крупный"
    MasterImportSub "Очаг1_Мелкий"
    MasterImportSub "Очаг2_Мелкий"
    MasterImportSub "Очаг3_Мелкий"
    MasterImportSub "Очаг4_Мелкий"
    MasterImportSub "Очаг1_Средний"
    MasterImportSub "Очаг2_Средний"
    MasterImportSub "Очаг3_Средний"
    MasterImportSub "Очаг4_Средний"
    MasterImportSub "Очаг1_Крупный"
    MasterImportSub "Очаг2_Крупный"
    MasterImportSub "Очаг3_Крупный"
    MasterImportSub "Очаг4_Крупный"
    MasterImportSub "Огненный шторм"
    MasterImportSub "Обрушение"

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
Dim v_Ctrl As CommandBarControl
'Dim SecExists As Boolean
    
'---Включаем обработку ошибок
    On Error GoTo Tail

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
    MsgBox "В ходе работы программы возникла ошибка! Если она будет повторяться - обратитесь к разработчику.", , ThisDocument.Name
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

Private Sub app_CellChanged(ByVal cell As IVCell)
'---Один раз в выполняем обновление иконок на кнопках
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEventFireArea.PictureRefresh
        ButEventStorm.PictureRefresh
        ButEventFog.PictureRefresh
        ButEventRush.PictureRefresh
        cellChangedCount = 0
    End If
End Sub


