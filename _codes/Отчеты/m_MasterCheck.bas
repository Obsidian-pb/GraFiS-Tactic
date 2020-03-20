Attribute VB_Name = "m_MasterCheck"
Option Explicit
'-------------------------------------Модуль для работы с формой MCheckForm и классом InfoCollector----------------------------
Dim remarks(27) As Boolean 'массив переменных для каждого из замечаний
Dim vOC_InfoAnalizer As InfoCollector
Public remarksHided As Integer 'переменная количества скрытых замечаний
'Public bo_GDZSRezRoundUp As Boolean 'округление резервных звеньев в большую сторону
'=================================== MASTER CHECK by Vasilchenko ================================================
Public Sub MasterCheckRefresh()
'Процедура реакции на действие пользователя
Dim i As Integer
Dim psi_TargetPageIndex As Integer
Dim Comment As Boolean

Dim strStvolCount As String
Dim strHoseCount As String

    psi_TargetPageIndex = Application.ActivePage.Index

'---обнуляем имеющиеся значения свойств
    If vOC_InfoAnalizer Is Nothing Then
        Set vOC_InfoAnalizer = New InfoCollector
    End If
    vOC_InfoAnalizer.sC_Refresh (psi_TargetPageIndex)

'---Очищаем форму и задаем стартовые условия
    MCheckForm.ListBox1.Clear
    MCheckForm.ListBox2.Clear
    
    Comment = False
    remarksHided = 0
    
'---Запускаем условия обработки
    'Ochag
    If remarks(0) = False Then
        If vOC_InfoAnalizer.pi_OchagCount = 0 Then
            If vOC_InfoAnalizer.pi_SmokeCount > 0 Or vOC_InfoAnalizer.pi_DevelopCount > 0 Or vOC_InfoAnalizer.pi_FireCount Then
                MCheckForm.ListBox1.AddItem "Не указан очаг пожара"
                Comment = True
            End If
        End If
    End If
    
    If remarks(1) = False Then
        If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_SmokeCount = 0 Then
            MCheckForm.ListBox1.AddItem "Не указаны зоны задымления"
            Comment = True
        End If
    End If
    
    If remarks(2) = False Then
        If vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 And vOC_InfoAnalizer.pi_DevelopCount = 0 Then
            MCheckForm.ListBox1.AddItem "Не указаны пути распространения пожара"
            Comment = True
        End If
    End If
    
    'Upravlenie
    If remarks(3) = False Then
        If vOC_InfoAnalizer.pi_BUCount >= 3 And vOC_InfoAnalizer.pi_ShtabCount = 0 Then
            MCheckForm.ListBox1.AddItem "Не создан оперативный штаб"
            Comment = True
        End If
    End If
    
    If remarks(4) = False Then
        If vOC_InfoAnalizer.pi_RNBDCount = 0 And vOC_InfoAnalizer.pi_OchagCount + vOC_InfoAnalizer.pi_FireCount > 0 Then
            MCheckForm.ListBox1.AddItem "Не указано решающее направление"
            Comment = True
        End If
    End If
    
    If remarks(5) = False Then
        If vOC_InfoAnalizer.pi_RNBDCount > 1 Then
            MCheckForm.ListBox1.AddItem "Решающее напраление должно быть одним"
            Comment = True
        End If
    End If
    
    If remarks(6) = False Then
        If vOC_InfoAnalizer.pi_BUCount >= 5 And vOC_InfoAnalizer.pi_SPRCount <= 1 Then
            MCheckForm.ListBox1.AddItem "Не организованы секторы проведения работ"
            Comment = True
        End If
    End If
    
    'GDZS
    If remarks(7) = False Then
        If vOC_InfoAnalizer.pi_GDZSpbCount < vOC_InfoAnalizer.pi_GDZSChainsCount Then
            MCheckForm.ListBox1.AddItem "Не выставлены посты безопасности для каждого звена ГДЗС (" & vOC_InfoAnalizer.pi_GDZSpbCount & "/" & vOC_InfoAnalizer.pi_GDZSChainsCount & ")"
            Comment = True
        End If
    End If
    
    If remarks(8) = False Then
        If vOC_InfoAnalizer.pi_GDZSChainsCount >= 3 And vOC_InfoAnalizer.pi_KPPCount = 0 Then
            MCheckForm.ListBox1.AddItem "Не создан контрольно-пропускной пункт ГДЗС"
            Comment = True
        End If
    End If
    
    If remarks(9) = False Then
        If vOC_InfoAnalizer.pb_GDZSDiscr = True Then
            MCheckForm.ListBox1.AddItem "В сложных условиях звенья ГДЗС должны состоять не менее чем из пяти газодымозащитников"
            Comment = True
        End If
    End If
    
    If remarks(10) = False Then
'        If Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) > vOC_InfoAnalizer.pi_GDZSChainsRezCount And bo_GDZSRezRoundUp = False Then
'                MCheckForm.ListBox1.AddItem "Недостаточно резервных звеньев ГДЗС с округлением в меньшую сторону (" & vOC_InfoAnalizer.pi_GDZSChainsRezCount & "/" & Fix(vOC_InfoAnalizer.ps_GDZSChainsRezNeed) & ")"
'                Comment = True
'        End If
    End If
    
    If remarks(10) = False Then
'        If Fix((vOC_InfoAnalizer.ps_GDZSChainsRezNeed / 0.3334) * 0.3333) + 1 > vOC_InfoAnalizer.pi_GDZSChainsRezCount And bo_GDZSRezRoundUp = True And vOC_InfoAnalizer.pi_GDZSChainsCount <> 0 Then
'                MCheckForm.ListBox1.AddItem "Недостаточно резервных звеньев ГДЗС с округлением в большую сторону (" & vOC_InfoAnalizer.pi_GDZSChainsRezCount & "/" & Fix((vOC_InfoAnalizer.ps_GDZSChainsRezNeed / 0.3334) * 0.3333) + 1 & ")"
'                Comment = True
'        End If
    End If
    
    'PPW
    If remarks(11) = False Then
        If vOC_InfoAnalizer.pi_WaterSourceCount > vOC_InfoAnalizer.pi_distanceCount Then
            MCheckForm.ListBox1.AddItem "Не указаны расстояния от каждого водоисточника до места пожара (" & vOC_InfoAnalizer.pi_distanceCount & "/" & vOC_InfoAnalizer.pi_WaterSourceCount & ")"
            Comment = True
        End If
    End If
    
    'Hoses
    If remarks(12) = False Then
        If vOC_InfoAnalizer.pi_WorklinesCount > vOC_InfoAnalizer.pi_linesPosCount Then
            MCheckForm.ListBox1.AddItem "Не указаны положения (этаж) для каждой рабочей линии (" & vOC_InfoAnalizer.pi_linesPosCount & "/" & vOC_InfoAnalizer.pi_WorklinesCount & ")"
            Comment = True
        End If
    End If
    
    If remarks(13) = False Then
        If vOC_InfoAnalizer.pi_linesCount > vOC_InfoAnalizer.pi_linesLableCount Then
            MCheckForm.ListBox1.AddItem "Не указаны диаметры для каждой рукавной линии (" & vOC_InfoAnalizer.pi_linesLableCount & "/" & vOC_InfoAnalizer.pi_linesCount & ")"
            Comment = True
        End If
    End If
    
    'План на местности
    If remarks(14) = False Then

        If vOC_InfoAnalizer.pi_BuildCount > vOC_InfoAnalizer.pi_SOCount Then
            MCheckForm.ListBox1.AddItem "Не указаны подписи степени огнестойкости для каждого из зданий (" & vOC_InfoAnalizer.pi_SOCount & "/" & vOC_InfoAnalizer.pi_BuildCount & ")"
            Comment = True
        End If
    End If
    
    If remarks(15) = False Then
        If vOC_InfoAnalizer.pi_OrientCount = 0 And vOC_InfoAnalizer.pi_BuildCount > 0 Then
            MCheckForm.ListBox1.AddItem "Не указаны ориентиры на местности, такие как роза ветров или подпись улицы"
            Comment = True
        End If
    End If
    
    'Показ расчетных данных
    If remarks(16) = False Then
        If vOC_InfoAnalizer.ps_FactStreemW <> 0 And vOC_InfoAnalizer.ps_FactStreemW < vOC_InfoAnalizer.ps_NeedStreemW Then
            MCheckForm.ListBox1.AddItem "Недостаточный фактический расход воды (" & vOC_InfoAnalizer.ps_FactStreemW & " л/c < " & vOC_InfoAnalizer.ps_NeedStreemW & " л/с)"
            Comment = True
        End If
    End If
    
    If remarks(17) = False Then
        If (vOC_InfoAnalizer.ps_FactStreemW * 600) > vOC_InfoAnalizer.pi_WaterValueHave Then
            If PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) > vOC_InfoAnalizer.pi_GetingWaterCount Then MCheckForm.ListBox1.AddItem "Недостаточное водоснабжение боевых позиций" '& (" & vOC_InfoAnalizer.pi_GetingWaterCount & "/" & PF_RoundUp(vOC_InfoAnalizer.ps_FactStreemW / 32) & ")"
            Comment = True
        End If
    End If
    
    If remarks(18) = False Then
        If vOC_InfoAnalizer.pi_PersonnelHave < vOC_InfoAnalizer.pi_PersonnelNeed Then
            MCheckForm.ListBox1.AddItem "Недостаточно личного состава, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_PersonnelHave & "/" & vOC_InfoAnalizer.pi_PersonnelNeed & ")"
            Comment = True
        End If
    End If
    
    If remarks(19) = False Then
        If vOC_InfoAnalizer.pi_Hoses51Have < vOC_InfoAnalizer.pi_Hoses51Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 51 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses51Have & "/" & vOC_InfoAnalizer.pi_Hoses51Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(20) = False Then
        If vOC_InfoAnalizer.pi_Hoses66Have < vOC_InfoAnalizer.pi_Hoses66Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 66 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses66Have & "/" & vOC_InfoAnalizer.pi_Hoses66Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(21) = False Then
        If vOC_InfoAnalizer.pi_Hoses77Have < vOC_InfoAnalizer.pi_Hoses77Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 77 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses77Have & "/" & vOC_InfoAnalizer.pi_Hoses77Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(22) = False Then
        If vOC_InfoAnalizer.pi_Hoses89Have < vOC_InfoAnalizer.pi_Hoses89Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 89 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses89Have & "/" & vOC_InfoAnalizer.pi_Hoses89Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(23) = False Then
        If vOC_InfoAnalizer.pi_Hoses110Have < vOC_InfoAnalizer.pi_Hoses110Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 110 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses110Have & "/" & vOC_InfoAnalizer.pi_Hoses110Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(24) = False Then
        If vOC_InfoAnalizer.pi_Hoses150Have < vOC_InfoAnalizer.pi_Hoses150Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 150 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses150Have & "/" & vOC_InfoAnalizer.pi_Hoses150Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(25) = False Then
        If vOC_InfoAnalizer.pi_Hoses200Have < vOC_InfoAnalizer.pi_Hoses200Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 200 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses200Have & "/" & vOC_InfoAnalizer.pi_Hoses200Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(26) = False Then
        If vOC_InfoAnalizer.pi_Hoses250Have < vOC_InfoAnalizer.pi_Hoses250Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 250 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses250Have & "/" & vOC_InfoAnalizer.pi_Hoses250Count & ")"
            Comment = True
        End If
    End If
    
    If remarks(27) = False Then
        If vOC_InfoAnalizer.pi_Hoses300Have < vOC_InfoAnalizer.pi_Hoses300Count Then
            MCheckForm.ListBox1.AddItem "Недостаточно напорных рукавов 300 мм, с учетом прибывшей техники (" & vOC_InfoAnalizer.pi_Hoses300Have & "/" & vOC_InfoAnalizer.pi_Hoses300Count & ")"
            Comment = True
        End If
    End If

    If Comment = False Then MCheckForm.ListBox1.AddItem "Замечаний не обнаружено"

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
        If vOC_InfoAnalizer.pi_StvolWBHave <> 0 Then strStvolCount = strStvolCount & vOC_InfoAnalizer.pi_StvolWBHave & " ств. ""Б"", "
        If vOC_InfoAnalizer.pi_StvolWAHave <> 0 Then strStvolCount = strStvolCount & vOC_InfoAnalizer.pi_StvolWAHave & " ств. ""А"", "
        If vOC_InfoAnalizer.pi_StvolWLHave <> 0 Then strStvolCount = strStvolCount & vOC_InfoAnalizer.pi_StvolWLHave & " лафетных, "
        strStvolCount = Left(strStvolCount, Len(strStvolCount) - 2)
        MCheckForm.ListBox2.AddItem "Водяных стволов"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_StvolWHave & " (" & strStvolCount & ")"
        strStvolCount = ""
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
    If vOC_InfoAnalizer.pi_HosesCount <> 0 Then
        If vOC_InfoAnalizer.pi_Hoses38Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses38Count & " - 38 мм, "
        If vOC_InfoAnalizer.pi_Hoses51Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses51Count & " - 51 мм, "
        If vOC_InfoAnalizer.pi_Hoses77Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses77Count & " - 77 мм, "
        If vOC_InfoAnalizer.pi_Hoses66Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses66Count & " - 66 мм, "
        If vOC_InfoAnalizer.pi_Hoses89Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses89Count & " - 89 мм, "
        If vOC_InfoAnalizer.pi_Hoses110Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses110Count & " - 110 мм, "
        If vOC_InfoAnalizer.pi_Hoses150Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses150Count & " - 150 мм, "
        If vOC_InfoAnalizer.pi_Hoses200Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses200Count & " - 200 мм, "
        If vOC_InfoAnalizer.pi_Hoses250Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses250Count & " - 250 мм, "
        If vOC_InfoAnalizer.pi_Hoses300Count <> 0 Then strHoseCount = strHoseCount & vOC_InfoAnalizer.pi_Hoses300Count & " - 300 мм, "
        strHoseCount = Left(strHoseCount, Len(strHoseCount) - 2)
        MCheckForm.ListBox2.AddItem "Задействовано напорных рукавов"
        MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_HosesCount & " (" & strHoseCount & ")"
        strHoseCount = ""
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

'    MCheckForm.ListBox2.AddItem "Имеется рукавов разных диаметров, с учетом прибывшей техники"
'    MCheckForm.ListBox2.List(MCheckForm.ListBox2.ListCount - 1, 1) = vOC_InfoAnalizer.pi_HosesHave

'Подсчет скрытых замечаний
    For i = 0 To UBound(remarks)
        If remarks(i) = True Then remarksHided = remarksHided + 1
    Next
        
End Sub

Public Sub RestoreComment()
'Обнуляем значения переменных не учитываемых замечений
Dim i As Integer
    
    For i = 0 To UBound(remarks)
        remarks(i) = False
    Next
    remarksHided = 0
End Sub

Public Sub HideComment()
'Скрываем замечания по желанию пользователя
    On Error Resume Next
    
    'И все равно мне не нравится как это выглядит надо  что-то другое придумать. Такой код чреват ошибками
    remarks(0) = MCheckForm.ListBox1.Value = "Не указан очаг пожара"
    remarks(1) = MCheckForm.ListBox1.Value = "Не указаны зоны задымления"
    remarks(2) = MCheckForm.ListBox1.Value = "Не указаны пути распространения пожара"
    remarks(3) = MCheckForm.ListBox1.Value = "Не создан оперативный штаб"
    remarks(4) = MCheckForm.ListBox1.Value = "Не указано решающее направление"
    remarks(5) = MCheckForm.ListBox1.Value = "Решающее напраление должно быть одним"
    remarks(6) = MCheckForm.ListBox1.Value = "Не организованы секторы проведения работ"
    remarks(7) = InStr(1, MCheckForm.ListBox1.Value, "Не выставлены посты безопасности") > 0
    remarks(8) = MCheckForm.ListBox1.Value = "Не создан контрольно-пропускной пункт ГДЗС"
    remarks(9) = InStr(1, MCheckForm.ListBox1.Value, "В сложных условиях звенья ГДЗС") > 0
    remarks(10) = InStr(1, MCheckForm.ListBox1.Value, "резервных звеньев") > 0
    remarks(11) = InStr(1, MCheckForm.ListBox1.Value, "расстояния от каждого") > 0
    remarks(12) = InStr(1, MCheckForm.ListBox1.Value, "положения") > 0
    remarks(13) = InStr(1, MCheckForm.ListBox1.Value, "диаметры") > 0
    remarks(14) = InStr(1, MCheckForm.ListBox1.Value, "подписи степени огнестойкости") > 0
    remarks(15) = InStr(1, MCheckForm.ListBox1.Value, "ориентиры на местности") > 0
    remarks(16) = InStr(1, MCheckForm.ListBox1.Value, "Недостаточный фактический расход") > 0
    remarks(17) = InStr(1, MCheckForm.ListBox1.Value, "Недостаточное водоснабжение") > 0
    remarks(18) = InStr(1, MCheckForm.ListBox1.Value, "Недостаточно личного состава") > 0
    remarks(19) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 51 мм,") > 0
    remarks(20) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 66 мм") > 0
    remarks(21) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 77 мм") > 0
    remarks(22) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 89 мм") > 0
    remarks(23) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 110 мм") > 0
    remarks(24) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 150 мм") > 0
    remarks(25) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 200 мм") > 0
    remarks(26) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 250 мм") > 0
    remarks(27) = InStr(1, MCheckForm.ListBox1.Value, "рукавов 300 мм") > 0
    
    On Error GoTo 0

    MasterCheckRefresh
End Sub




