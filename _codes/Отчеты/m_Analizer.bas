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
    For i = 1 To vsO_BaseShape.Shapes.Count
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
        fp_SetValue = PF_RoundUp(vOC_InfoAnalizer.ps_GDZSChainsRezCount)
    Case Is = "Требуется звеньев ГДЗС"
'        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.pi_GDZSChainsCount / 3)
        fp_SetValue = vOC_InfoAnalizer.pi_GDZSChainsCount + PF_RoundUp(vOC_InfoAnalizer.ps_GDZSChainsRezCount)
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

'=================================== MASTER CHECK by Vasilchenko ================================================
Public Sub MasterCheckRefresh()
'Процедура реакции на действие пользователя
Dim i As Integer
Dim psi_TargetPageIndex As Integer

    psi_TargetPageIndex = Application.ActivePage.Index

'---обнуляем имеющиеся значения свойств
'    psi_TargetPageIndex = ActivePage.Index
    If vOC_InfoAnalizer Is Nothing Then
        Set vOC_InfoAnalizer = New InfoCollector
    End If
    vOC_InfoAnalizer.sC_Refresh (psi_TargetPageIndex)

'---Запускаем условия обработки
    MCheckForm.ListBox1.Clear
    MCheckForm.ListBox2.Clear
    Dim comment As Boolean
    comment = False
    'Ochag
    If vOC_InfoAnalizer.pi_OchagCount = 0 Then
        If vOC_InfoAnalizer.pi_SmokeCount > 0 Or vOC_InfoAnalizer.pi_DevelopCount > 0 Or vOC_InfoAnalizer.pi_FireCount Then
            MCheckForm.ListBox1.AddItem "Не указан очаг пожара"
            comment = True
        End If
    End If
    If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_SmokeCount = 0 Then
        MCheckForm.ListBox1.AddItem "Не указаны зоны задымления"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_DevelopCount = 0 Then
        MCheckForm.ListBox1.AddItem "Не указаны пути распространения пожара"
        comment = True
    End If
    'Upravlenie
    If vOC_InfoAnalizer.pi_BUCount >= 3 And vOC_InfoAnalizer.pi_ShtabCount = 0 Then
        MCheckForm.ListBox1.AddItem "Не создан оперативный штаб"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_RNBDCount = 0 And vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 Then
        MCheckForm.ListBox1.AddItem "Не указано решающее направление"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_RNBDCount > 1 Then
        MCheckForm.ListBox1.AddItem "Решающее напраление должно быть одним"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_BUCount >= 5 And vOC_InfoAnalizer.pi_SPRCount <= 1 Then
        MCheckForm.ListBox1.AddItem "Не организованы секторы проведения работ"
        comment = True
    End If
    'GDZS
    If vOC_InfoAnalizer.pi_GDZSpbCount < vOC_InfoAnalizer.pi_GDZSChainsCount Then
        MCheckForm.ListBox1.AddItem "Не выставлены посты безопасности для каждого звена ГДЗС (" & vOC_InfoAnalizer.pi_GDZSpbCount & "/" & vOC_InfoAnalizer.pi_GDZSChainsCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsCount >= 3 And vOC_InfoAnalizer.pi_KPPCount = 0 Then
        MCheckForm.ListBox1.AddItem "Не создан контрольно-пропускной пункт ГДЗС"
        comment = True
    End If
    If vOC_InfoAnalizer.pb_GDZSDiscr = True Then
        MCheckForm.ListBox1.AddItem "В сложных условиях звенья ГДЗС должны состоять не менее чем из пяти газодымозащитников"
        comment = True
    End If
    If Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) - vOC_InfoAnalizer.pi_GDZSChainsRezCount <> 0 Then
        MCheckForm.ListBox1.AddItem "Недостаточно резервных звеньев ГДЗС (" & vOC_InfoAnalizer.pi_GDZSChainsRezCount & "/" & Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) & ")"
        comment = True
    End If
    'PPW
    If vOC_InfoAnalizer.pi_WaterSourceCount > vOC_InfoAnalizer.pi_distanceCount Then
        MCheckForm.ListBox1.AddItem "Не указаны расстояния от каждого водоисточника до места пожара (" & vOC_InfoAnalizer.pi_distanceCount & "/" & vOC_InfoAnalizer.pi_WaterSourceCount & ")"
        comment = True
    End If
    'Hoses
'    If vOC_InfoAnalizer.pb_AllHosesWithPos Then MCheckForm.ListBox1.AddItem "Не указаны положения (этаж) для каждой рабочей линии"
    If vOC_InfoAnalizer.pi_WorklinesCount > vOC_InfoAnalizer.pi_linesPosCount Then
        MCheckForm.ListBox1.AddItem "Не указаны положения (этаж) для каждой рабочей линии (" & vOC_InfoAnalizer.pi_linesPosCount & "/" & vOC_InfoAnalizer.pi_WorklinesCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_linesCount > vOC_InfoAnalizer.pi_linesLableCount Then
        MCheckForm.ListBox1.AddItem "Не указаны диаметры для каждой рукавной линии (" & vOC_InfoAnalizer.pi_linesLableCount & "/" & vOC_InfoAnalizer.pi_linesCount & ")"
        comment = True
    End If
    'Plan na mestnosti
    If vOC_InfoAnalizer.pi_BuildCount > vOC_InfoAnalizer.pi_SOCount Then
        MCheckForm.ListBox1.AddItem "Не указаны подписи степени огнестойкости для каждого из зданий (" & vOC_InfoAnalizer.pi_SOCount & "/" & vOC_InfoAnalizer.pi_BuildCount & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_OrientCount = 0 And vOC_InfoAnalizer.pi_BuildCount > 0 Then
        MCheckForm.ListBox1.AddItem "Не указаны ориентиры на местности, такие как роза ветров или подпись улицы"
        comment = True
    End If
            'показ расчетных данных
         
    If vOC_InfoAnalizer.ps_FactStreemW <> 0 And vOC_InfoAnalizer.ps_FactStreemW < vOC_InfoAnalizer.ps_NeedStreemW Then
        MCheckForm.ListBox1.AddItem "Недостаточный фактический расход воды (" & vOC_InfoAnalizer.ps_FactStreemW & " л/c < " & vOC_InfoAnalizer.ps_NeedStreemW & " л/с)"
        comment = True
    End If
    If (vOC_InfoAnalizer.ps_FactStreemW * 600) > vOC_InfoAnalizer.pi_WaterValueHave Then
        If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) > vOC_InfoAnalizer.pi_GetingWaterCount Then MCheckForm.ListBox1.AddItem "Недостаточное водоснабжение боевых позиций" '& (" & vOC_InfoAnalizer.pi_GetingWaterCount & "/" & PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) & ")"
        comment = True
    End If
    If vOC_InfoAnalizer.pi_PersonnelHave < vOC_InfoAnalizer.pi_PersonnelNeed Then
        MCheckForm.ListBox1.AddItem "Недостаточно личного состава, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_PersonnelHave & "/" & vOC_InfoAnalizer.pi_PersonnelNeed & ")"
        comment = True
    End If
    If comment = False Then MCheckForm.ListBox1.AddItem "Замечаний не обнаружено"

    '============Вторая вкладка - Сводка тактических данных===========
    If vOC_InfoAnalizer.pi_TechTotalCount <> 0 Then
        MCheckForm.ListBox2.AddItem "Техники РСЧС"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_TechTotalCount
    End If
    If vOC_InfoAnalizer.pi_MVDCount <> 0 Then
        MCheckForm.ListBox2.AddItem "Техники МВД"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MVDCount
    End If
    If vOC_InfoAnalizer.pi_MZdravCount <> 0 Then
        MCheckForm.ListBox2.AddItem "Техники Минздрав"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MZdravCount
    End If
    If vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount <> 0 Then
        MCheckForm.ListBox2.AddItem "Техники иных ведомств"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_TechTotalCount - vOC_InfoAnalizer.pi_FireTotalCount - vOC_InfoAnalizer.pi_MVDCount - vOC_InfoAnalizer.pi_MZdravCount
    End If
    If vOC_InfoAnalizer.pi_FireTotalCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Техники пожарной охраны"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_FireTotalCount
    End If
    If vOC_InfoAnalizer.pi_TechTotalCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Основных ПА"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_MainPAHave + vOC_InfoAnalizer.pi_TargetedPAHave & " (" & vOC_InfoAnalizer.pi_TargetedPAHave & " цел.прим., " & vOC_InfoAnalizer.pi_MainPAHave & " общ.прим.)"
    End If
    If vOC_InfoAnalizer.pi_GetingWaterCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Задействовано водоисточников"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GetingWaterCount
    End If
    If vOC_InfoAnalizer.pi_SpecialPAHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Специальных ПА"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_SpecialPAHave & " (" & vOC_InfoAnalizer.pi_ALCount + vOC_InfoAnalizer.pi_AKPCount & " высотных)"
    End If
    If vOC_InfoAnalizer.pi_OtherTechincsHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Прочей пож.техники"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_OtherTechincsHave
    End If
    If vOC_InfoAnalizer.pi_BUCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Боевых участков"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_BUCount
    End If
    If vOC_InfoAnalizer.pi_SPRCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Секторов проведения работ"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_SPRCount
    End If
    If vOC_InfoAnalizer.pi_PersonnelHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Личного состава (без водителей)"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_PersonnelHave
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Работает звеньев ГДЗС"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GDZSChainsCount & " (" & vOC_InfoAnalizer.pi_GDZSMansCount & " газодымозащитников)"
    End If
    If vOC_InfoAnalizer.pi_GDZSChainsRezCount <> 0 Then
         MCheckForm.ListBox2.AddItem "Звеньев ГДЗС в резерве"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_GDZSChainsRezCount & " (" & vOC_InfoAnalizer.pi_GDZSMansRezCount & " газодымозащитников)"
    End If
    If vOC_InfoAnalizer.ps_FireSquare <> 0 Then
         MCheckForm.ListBox2.AddItem "Площадь пожара"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_FireSquare & " м кв. (плащадь тушения " & vOC_InfoAnalizer.ps_ExtSquare & " м кв.)"
    End If
    If vOC_InfoAnalizer.pi_StvolWHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Водяных стволов"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolWHave & " (" & vOC_InfoAnalizer.pi_StvolWBHave & " ств.Б, " & vOC_InfoAnalizer.pi_StvolWAHave & " ств.А, " & vOC_InfoAnalizer.pi_StvolWLHave & " лафетных)"
    End If
    If vOC_InfoAnalizer.pi_StvolFoamHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Пенных стволов"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolFoamHave
    End If
    If vOC_InfoAnalizer.pi_StvolPowderHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Порошковых стволов"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolPowderHave
    End If
    If vOC_InfoAnalizer.pi_StvolGasHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Подано газовых стволов"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolGasHave
    End If
    If vOC_InfoAnalizer.ps_FactStreemW <> 0 Then
         MCheckForm.ListBox2.AddItem "Фактический расход воды"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_FactStreemW & " л/с"
           
    End If
    If vOC_InfoAnalizer.ps_NeedStreemW <> 0 Then
         MCheckForm.ListBox2.AddItem "Требуемый расход воды"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_NeedStreemW & " л/с"
    End If
    If vOC_InfoAnalizer.pi_WaterValueHave <> 0 Then
         MCheckForm.ListBox2.AddItem "Запас воды в емкостях ПА"
         MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_WaterValueHave / 1000 & " т"
    End If
    If vOC_InfoAnalizer.pi_linesCount - vOC_InfoAnalizer.pi_WorklinesCount <> 0 Then
        MCheckForm.ListBox2.AddItem "Магистральных линий"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_linesCount - vOC_InfoAnalizer.pi_WorklinesCount
    End If
    If vOC_InfoAnalizer.pi_HosesLength <> 0 Then
        MCheckForm.ListBox2.AddItem "Общая длина напорных линий"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_HosesLength & " м"
    End If
    If vOC_InfoAnalizer.pi_Hoses51Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 51 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses51Count
    End If
    If vOC_InfoAnalizer.pi_Hoses66Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 66 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses66Count
    End If
    If vOC_InfoAnalizer.pi_Hoses77Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 77 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses77Count
    End If
    If vOC_InfoAnalizer.pi_Hoses89Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 89 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses89Count
    End If
    If vOC_InfoAnalizer.pi_Hoses110Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 110 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses110Count
    End If
    If vOC_InfoAnalizer.pi_Hoses150Count <> 0 Then
        MCheckForm.ListBox2.AddItem "Количество рукавов 150 мм"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_Hoses150Count
    End If
    If vOC_InfoAnalizer.ps_GetedWaterValue <> 0 Then
        MCheckForm.ListBox2.AddItem "Забирается воды"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.ps_GetedWaterValue & " л/с (" & "max = " & vOC_InfoAnalizer.ps_GetedWaterValueMax & " л/с)"
    End If
    If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) <> 0 Then
        MCheckForm.ListBox2.AddItem "Требуется установить ПН-40 на ИНППВ (по расходу)"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32)
    End If
    If vOC_InfoAnalizer.ps_FactStreemW * 600 <> 0 Then
        MCheckForm.ListBox2.AddItem "Требуемый запас воды (по расходу, на 10 мин)"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = (vOC_InfoAnalizer.ps_FactStreemW * 600) / 1000 & " т"
    End If
        
End Sub
