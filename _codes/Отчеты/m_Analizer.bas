Attribute VB_Name = "m_Analizer"
Option Explicit
'--------------------------------------------------Модуль для работы с классом InfoCollector----------------------------
Dim vOC_InfoAnalizer As InfoCollector



Public Sub sP_InfoCollectorActivate()
    Set vOC_InfoAnalizer = New InfoCollector
End Sub

Public Sub sP_InfoCollectorDeActivate()
    Set vOC_InfoAnalizer = Nothing
End Sub

Public Sub sP_ChangeValue(ShpObj As Visio.Shape)
'Процедура реакции на действие пользователя
Dim i As Integer
Dim psi_TargetPageIndex As Integer

'---Предлагаем пользователю указать страницу для анализа
    SeetsSelectForm.Show
    psi_TargetPageIndex = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet).Index

'---обнуляем имеющиеся значения свойств
'    psi_TargetPageIndex = ActivePage.Index
    If vOC_InfoAnalizer Is Nothing Then
        Set vOC_InfoAnalizer = New InfoCollector
    End If
    vOC_InfoAnalizer.sC_Refresh (psi_TargetPageIndex)

'---Запускаем циклы обработки фигур отчета
    sP_ChangeValueMain ShpObj.ID, psi_TargetPageIndex

End Sub

Public Sub sP_ChangeValueMain(asi_ShpInd As Integer, asi_TargetPage As Integer)
'Процедура реакции на действие пользователя
Dim i As Integer
Dim vsO_TargetPage As Visio.Page
Dim vsO_BaseShape As Visio.Shape
Dim vsO_Shape As Visio.Shape

    On Error GoTo EX
'---Определяем базовую фигуру
    Set vsO_TargetPage = Application.ActiveDocument.Pages(asi_TargetPage)
    Set vsO_BaseShape = Application.ActivePage.Shapes.ItemFromID(asi_ShpInd)

'---Перебираем все фигуры в отчете
    For i = 1 To vsO_BaseShape.Shapes.count
        Set vsO_Shape = vsO_BaseShape.Shapes(i)
        If vsO_Shape.CellExists("Actions.ChangeValue", 0) = True Then '---Проверяем является ли фигура ОТЧЕТОМ!!!
            sP_ChangeValueMain vsO_Shape.ID, asi_TargetPage
        End If
        If vsO_Shape.CellExists("User.PropertyValue", 0) = True Then '---Проверяем является ли фигура полем отчета
            vsO_Shape.Cells("User.PropertyValue").FormulaU = _
                str(fp_SetValue(vsO_Shape.Cells("Prop.PropertyName").ResultStr(visUnitsString)))
        End If
    Next i

Set vsO_Shape = Nothing
Set vsO_BaseShape = Nothing
Exit Sub
EX:
    Set vsO_Shape = Nothing
    Set vsO_BaseShape = Nothing
    SaveLog Err, "sP_ChangeValueMain"
End Sub



Private Function fp_SetValue(ass_PropertyName As String) As Double
'Процедура устанавливает в поле значения фигуры значение свойства
'---Проверяем значение какого именно свойства требется получить
'fp_SetValue = 111
Select Case ass_PropertyName
    
    '---НОВЫЕ-------------
    Case Is = "Основных ПА"
        fp_SetValue = vOC_InfoAnalizer.pi_MainPAHave + vOC_InfoAnalizer.pi_TargetedPAHave
    Case Is = "Высотных ПА"
        fp_SetValue = vOC_InfoAnalizer.pi_ALCount + vOC_InfoAnalizer.pi_AKPCount
    Case Is = "Имеется АЦ"
        fp_SetValue = vOC_InfoAnalizer.pi_ACCount
    Case Is = "Имеется АГДЗС"
        fp_SetValue = vOC_InfoAnalizer.pi_AGCount
    Case Is = "Имеется автолестниц"
        fp_SetValue = vOC_InfoAnalizer.pi_ALCount
    Case Is = "Имеется автоподъемников"
        fp_SetValue = vOC_InfoAnalizer.pi_AKPCount
    Case Is = "Имеется техники МВД"
        fp_SetValue = vOC_InfoAnalizer.pi_MVDCount
    Case Is = "Имеется техники Минздрав"
        fp_SetValue = vOC_InfoAnalizer.pi_MZdravCount
    Case Is = "Боевых участков"
        fp_SetValue = vOC_InfoAnalizer.pi_BUCount
    Case Is = "Техники РСЧС"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount
    Case Is = "Техники пожарной охраны"
        fp_SetValue = vOC_InfoAnalizer.pi_FireTotalCount
    Case Is = "Техники не МЧС"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount
    Case Is = "Техники не МЧС (прочяя)"
        fp_SetValue = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount
    
    '---НОВЫЕ-------------
    
    
    Case Is = "Основных ПА общего назначения"
        fp_SetValue = vOC_InfoAnalizer.pi_MainPAHave
    Case Is = "Требуется АЦ"
        fp_SetValue = PF_RoundUp((vOC_InfoAnalizer.pi_PersonnelNeed + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3) * 3) / 4)
    Case Is = "Требуется АНР"
        fp_SetValue = PF_RoundUp((vOC_InfoAnalizer.pi_PersonnelNeed + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3) * 3) / 5)
    Case Is = "Основных ПА целевого назначения"
        fp_SetValue = vOC_InfoAnalizer.pi_TargetedPAHave
    Case Is = "Специальных ПА"
        fp_SetValue = vOC_InfoAnalizer.pi_SpecialPAHave
    Case Is = "Прочей техники"
        fp_SetValue = vOC_InfoAnalizer.pi_OtherTechincsHave
    Case Is = "Имеется личного состава"
        fp_SetValue = vOC_InfoAnalizer.pi_PersonnelHave
        
    Case Is = "Требуется личного состава" 'С учетом резервных звеньев
        fp_SetValue = vOC_InfoAnalizer.pi_PersonnelNeed + vOC_InfoAnalizer.pi_GDZSMansRezCount
'        + Int(vOC_InfoAnalizer.pi_HosesCount * 20 / 100)  ' РУКАВНЫЕ ЛИНИИ!!!!!
            
    Case Is = "Фактическое количество звеньев ГДЗС"
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount
    Case Is = "Требуется резервных звеньев"
'        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3)
'        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsRezCount)
        fp_SetValue = vOC_InfoAnalizer.ps_GDZSChainsRezNeed
    Case Is = "Требуется звеньев ГДЗС"
'        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3)
'        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsRezCount)
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + vOC_InfoAnalizer.ps_GDZSChainsRezNeed
    Case Is = "Фактическое количество газодымозащитников"
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSMansCount
    Case Is = "Площадь пожара"
        fp_SetValue = vOC_InfoAnalizer.ps_FireSquare
    Case Is = "Площадь тушения"
        fp_SetValue = vOC_InfoAnalizer.ps_ExtSquare
    Case Is = "Требуемый расход воды"
        fp_SetValue = vOC_InfoAnalizer.ps_NeedStreemW
    Case Is = "Фактический расход воды"
        fp_SetValue = vOC_InfoAnalizer.ps_FactStreemW
    Case Is = "Подано водяных стволов"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWHave
    Case Is = "Подано пенных стволов"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolFoamHave
    Case Is = "Подано порошковых стволов"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolPowderHave
    Case Is = "Подано газовых стволов"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolGasHave
    Case Is = "Подано стволов А"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWAHave
    Case Is = "Подано стволов Б"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWBHave
    Case Is = "Подано лафетных стволов"
        fp_SetValue = vOC_InfoAnalizer.pi_StvolWLHave


    Case Is = "Требуется подать стволов Б"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 3.7)
    Case Is = "Требуется подать стволов А"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 7.4)
    Case Is = "Требуется подать лафетных стволов"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_NeedStreemW / 12)
        
    Case Is = "Забирается воды"
        fp_SetValue = vOC_InfoAnalizer.ps_GetedWaterValue
    Case Is = "Максимум забираемой воды"
        fp_SetValue = vOC_InfoAnalizer.ps_GetedWaterValueMax
    Case Is = "Установлено на водоисточники"
        fp_SetValue = vOC_InfoAnalizer.pi_GetingWaterCount
    Case Is = "Требуется установить на водоисточники АЦ"
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32)

    Case Is = "Количество рукавов 51мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses51Count
    Case Is = "Количество рукавов 66мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses66Count
    Case Is = "Количество рукавов 77мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses77Count
    Case Is = "Количество рукавов 89мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses89Count
    Case Is = "Количество рукавов 110мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses110Count
    Case Is = "Количество рукавов 150мм"
        fp_SetValue = vOC_InfoAnalizer.pi_Hoses150Count
    Case Is = "Общая длина напорных линий"
        fp_SetValue = vOC_InfoAnalizer.pi_HosesLength
        
    Case Is = "Фактический запас воды"
        fp_SetValue = vOC_InfoAnalizer.pi_WaterValueHave
    Case Is = "Требуемый запас воды (10мин)"
        fp_SetValue = vOC_InfoAnalizer.ps_FactStreemW * 600   'Для 10 минут

End Select


End Function

