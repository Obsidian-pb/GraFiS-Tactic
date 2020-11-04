VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ButEvent As Class1
Dim C_ConnectionsTrace As c_HoseConnector
Dim WithEvents LineAppEvents As Visio.Application
Attribute LineAppEvents.VB_VarHelpID = -1


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'Процедура закрытия документа и удаления его рабочих элементов

'---Очищаем объект ButEvent и удаляем кнопку "Рукав" с панели управления "Превращения"
    Set ButEvent = Nothing
    DeleteButtonLine
    DeleteButtonMLine
    DeleteButtonVHose
    DeleteButtonNormalize
    DeleteButtonNRSSettings
    DeleteButtonNRSReport
    
'---В случае, если на панели "Превращения нет ни одной кнопки, удаляем её
    If Application.CommandBars("Превращения").Controls.Count = 0 Then RemoveTBImagination
    
'---Очищаем переменную приложения
Set LineAppEvents = Nothing
Set C_ConnectionsTrace = Nothing
End Sub


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo ex
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set LineAppEvents = Visio.Application
'---Активируем экземпляр класса для отслеживания соединений
    Set C_ConnectionsTrace = New c_HoseConnector

'---Импортируем мастера
    sp_MastersImport

'---Добавляем ячейку Аспект (если еще не была добавлена)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Создаем панель управления "Превращения" и добавляем на нее кнопку "Обратить в линию"
    AddTBImagination        'Добавляем тулбокс Обрашщения
    AddButtonLine           'Добавляем кнопку рабочей линии
    AddButtonMLine          'Добавляем фигуру магистральной линии
    AddButtonVHose          'Добавляем кнопку всасывающей линии
    AddButtonNormalize      'Добавляем кнопку "Расчет НРС"
    AddButtonNRSSettings    'Добавляем кнопку "Настройки расчета НРС"
    AddButtonNRSReport      'Добавляем кнопку "Отчет расчета НРС"

Set ButEvent = New Class1

'---Активируем опцию приклеивания к контурам фигур
    Application.ActiveDocument.GlueSettings = visGlueToGeometry + visGlueToGuides + visGlueToConnectionPoints

'---Добавляем для документа своство "FireTime"
    sm_AddFireTime
    
'---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates
    
Exit Sub
ex:
    SaveLog Err, "Document_DocumentOpened"
End Sub


Private Sub LineAppEvents_CellChanged(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer
'---Проверяем имя ячейки
    
    If cell.Name = "Prop.HoseMaterial" Or cell.Name = "Prop.HoseDiameter" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
        HoseDiametersListImport (ShpInd)
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
        HoseResistanceValueImport (ShpInd)
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
        HoseMaxFlowValueImport (ShpInd)
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
        HoseWeightValueImport (ShpInd)
    End If
    
    
    
'MsgBox Cell.Shape.Index
'MsgBox Cell.Shape.ID

'В случае, если произошло изменение не нужной ячейки прекращаем событие
End Sub

Private Sub sp_TemplatesImport()
    Visio.Application.ActivePage.Drop ThisDocument.Masters("ВсасывающийРукав"), 0, 0
End Sub


Private Sub sp_MastersImport()
'---Импортируем мастера

    MasterImportSub "Линии НРС.vss", "ВсасывающийРукав"
    MasterImportSub "Линии НРС.vss", "ВсасывающийРукав1000"
    MasterImportSub "Линии НРС.vss", "НапорноВсасывающийРукав"
    
End Sub


Private Sub sm_AddFireTime()
'Процедура добавляет в документ свойство - время начала пожара, в случае его отсутствия

    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If

End Sub

Private Sub LineAppEvents_ShapeAdded(ByVal Shape As IVShape)
'Событие добавления на лист фигуры
Dim v_Cntrl As CommandBarControl
Dim SecExists As Boolean
    
'---Включаем обработку ошибок
    On Error GoTo Tail

'---Проверяем какая кнопка нажата и в зависимости от этого выполняем действие
    For Each v_Ctrl In Application.CommandBars("Превращения").Controls
        If v_Ctrl.State = msoButtonDown Then
            Select Case v_Ctrl.Caption
                Case Is = "Рукав"
                    '---Запускаем процедуру обращения в рабочую рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру рабочей рукавной линии
'                            ButEvent.MakeHoseLine
                            MakeHoseLine 51, 0
                        End If
                    End If
                Case Is = "Всасывающий рукав"
                    '---Запускаем процедуру обращения во всасывающую рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру всасывающей рукавной линии
'                            ButEvent.MakeVHoseLine
                            MakeVHoseLine
                        End If
                    End If
                Case Is = "Магистральная линия"
                    '---Запускаем процедуру обращения в магистральную рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру магистральной рукавной линии
'                            ButEvent.MakeMagHoseLine
                            MakeHoseLine 77, 1
                        End If
                    End If
            End Select
        End If
    Next v_Ctrl
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчику."
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
