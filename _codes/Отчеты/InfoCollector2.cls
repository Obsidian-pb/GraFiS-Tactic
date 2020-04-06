VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoCollector2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------Класс сбора, обобщения и обобщения информации имеющейся на листе------------------------------------------



'----------------------------------------Коллекция элементов---------------------------------------------------------
Public elements As ElementsShell
'----------------------------------------Публичные свойства (настройки)----------------------------------------------
Public options As InfoCollectorOptions          'Настройки расчета

'----------------------------------------Внутренине свойства (служебные)---------------------------------------------

'----------------------------------------События класса---------------------------------------------------------
Public Event AfterRefresh()


'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
    Set elements = New ElementsShell
    Set options = New InfoCollectorOptions
End Sub
Private Sub Class_Terminate()
    Set elements = Nothing
    Set options = Nothing
End Sub




Public Function Refresh(a_PageIndex As Integer) As InfoCollector2
'Процедура обновления свойств класса
'Перебирает все элементы находящиеся на указанном a_PageIndex листе
Dim shp As Visio.Shape
Dim shpIndex As indexPers
Dim targetPage As Visio.Page
Dim tmp As Variant
Dim latestSquareTime As Date

    On Error GoTo EX
'Обнуляем количественные свойства
'---Обнуляем минимальное время площади пожара
    latestSquareTime = 0
'---Обнуляем прочие количественные свойства
    elements.Clear

'Устанавливаем индекс страницы для анализа
    Set targetPage = Application.ActiveDocument.Pages(a_PageIndex)

'---Перебираем все фигуры и в случае если фигура является фигурой ГраФиС анализируем её
    For Each shp In targetPage.Shapes
        If IsGFSShape(shp) Then         'Является ли фигура фигурой ГраФиС
'            If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then  'Проверяем версию фигуры
                shpIndex = shp.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
                
                With elements
                Select Case shpIndex
                '---Пожарные автомобили-----------
                    Case Is = ipAC   'Автоцистерны
                        .Raise "MainOverallHave;MainPAHave;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipANR 'АНР
                        .Raise "MainOverallHave;MainPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAPP 'АПП
                        .Raise "MainOverallHave;MainPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipAL 'АЛ
                        .Raise "SpecialPAHave;ALCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAKP 'АКП
                        .Raise "SpecialPAHave;AKPCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipASO 'АСО
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAT 'АТ
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipADU 'АДУ
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPNS 'ПНС
                        .Raise "TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAA 'AA
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipAV 'АВ
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAKT 'АКТ
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipAP 'АП
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAGVT 'АГВТ
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAG 'АГ
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAGDZS 'АГДЗС
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPNS 'ПКС
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPozhLab 'ЛБ
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipASA 'АСА
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipASH 'АШ
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAR 'АР
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
                        .Raise "TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        '>>>Учет возможности отсутствия ячеек "Prop.Dest" для старых фигур
                        tmp = CellVal(shp, "Prop.Dest", visUnitsString)
                        If tmp = "Специальный" _
                            Then .Raise "SpecialPAHave" ' pi_SpecialPAHave = pi_SpecialPAHave + 1
                        If InStr(1, tmp, "общ") <> 0 _
                            Then .Raise "MainPAHave" ' pi_MainPAHave = pi_MainPAHave + 1
                        If InStr(1, tmp, "целев") <> 0 _
                            Then .Raise "TargetedPAHave" ' pi_TargetedPAHave = pi_TargetedPAHave + 1
                        If InStr(1, tmp, "Основ") <> 0 _
                            Then .Raise "MainOverallHave" 'pi_MainOverallCount = pi_MainOverallCount + 1
                        '<<<Учет возможности отсутствия ячейки "Prop.Dest" для старых фигур
                    Case Is = indexPers.ipAvtoObschee 'Общее обозначение
                        If CellVal(shp, "Actions.RSCHS.Checked") = 1 _
                            Then .Raise "TotalTechCount" ' pi_TechTotalCount = pi_TechTotalCount + 1
                    Case Is = indexPers.ipACL 'АЦЛ
                        .Raise "MainOverallHave;MainPAHave;ALCount;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipACKP 'АЦКП
                        .Raise "MainOverallHave;MainPAHave;AKPCount;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                
                
                '---Прочая пожарная техника----------
                    Case Is = indexPers.ipPoezd 'Поезда
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                    Case Is = indexPers.ipSamolet 'Самолет
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipVertolet 'Вертолет
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipMotoPump 'Мотопомпы
                        .Raise "OtherTechincsHave;TotalFireCount"
                        .Raise "PersonnelNeed", shp.Cells("Prop.Personnel")
                    Case Is = indexPers.ipKorabl 'Корабль
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipKater 'Катер
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipTank 'Танки
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipMVD 'МВД
                        .Raise "TotalTechCount;MVDCount"
                    Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
                        .Raise "TotalTechCount;SMPCount"
                        
                '---Объекты пожара-----------------
                    Case Is = indexPers.ipPloschadPozhara 'Площадь пожара
                        .Raise "FireCount"
                        If CellVal(shp, "Prop.SquareTime", visDate) > latestSquareTime Then 'Проверяем является ли данная фигура последней по времени
                            .SetValByCellID shp, "User.FireSquare;User.ExtSquare;User.NeedStreamW"
                                                
                            latestSquareTime = shp.Cells("Prop.SquareTime") 'Устанавливаем новое значение наибольшего времени учета для площади
                        ElseIf shp.Cells("Prop.SquareTime") = latestSquareTime Then
                            .RaiseByCellID shp, "User.FireSquare;User.ExtSquare;User.NeedStreamW"
                        End If
                        
                    Case Is = indexPers.ipZadiml 'Задымление
                        .Raise "SmokeCount"
                    Case Is = indexPers.ipOchag 'Очаг
                        .Raise "OchagCount"
                    Case Is = indexPers.ipPostradavshie 'Гибель
                    Case Is = indexPers.ipRasprostran 'Распространение
                        .Raise "SpreadCount"
                    Case Is = indexPers.ipStorm 'Огненный шторм
                    Case Is = indexPers.ipOdinPozhar 'Одиночный пожар
                    Case Is = indexPers.ipPodpisStepenOgnestoykosty 'stepen ognestoikosti
                        .Raise "SOCount"
                
                '---Связь и освещение------------------
                    Case Is = indexPers.ipProzhektor 'Прожектор
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipNosimoeRadio 'Радиостанция носимая
                    Case Is = indexPers.ipAvtoRadio 'Радиостанция возимая
                    Case Is = indexPers.ipStacRadio 'Радиостанция стационарная
                    Case Is = indexPers.ipGromkogovoritel 'Громкоговоритель
                    Case Is = indexPers.ipTelefon 'Телефон
        
                '---ГДЗС----------------------------
                    Case Is = indexPers.ipGDZSZvenoVozd 'Звено - ДАСВ
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
                            .Raise "PersonnelNeed;GDZSMansCountWork", CellVal(shp, "Prop.Personnel")
                            
                            If shp.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                                .Raise "GDZSChainsCountWorkSoft"
                            Else
                                .Raise "GDZSChainsCountWorkHard"
                                
                                If CellVal(shp, "Prop.Personnel") < 5 Then .SetVal "GDZSDiscr", True
                            End If
                            'Вычисляем требуемое количество звеньев и людей
                            .SetVal "GDZSChainsCountNeed", .ByID("GDZSChainsCountWork").Result + .ByID("GDZSChainsRezCountNeed").Result
                            .SetVal "GDZSMansCountNeed", .ByID("GDZSMansCountWork").Result + .ByID("GDZSMansRezCountNeed").Result
                        End If
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "В резерве" Then
                            'ПРОВЕРИТЬ!!! Получается, что дважды считается требуемое количество людей
                            .Raise "GDZSMansRezCountHave", CellVal(shp, "Prop.Personnel")
                            .Raise "GDZSChainsRezCountHave"
                        End If
                        
                        
                    Case Is = indexPers.ipGDZSZvenoKisl 'Звено - ДАСК
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
                            .Raise "PersonnelNeed;GDZSMansCountWork", CellVal(shp, "Prop.Personnel")
                            If shp.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                                .Raise "GDZSChainsCountWorkSoft"
                            Else
                                .Raise "GDZSChainsCountWorkHard"
                                
                                If CellVal(shp, "Prop.Personnel") < 5 Then .SetVal "GDZSDiscr", True
                            End If
                            'Вычисляем требуемое количество звеньев и людей
                            .SetVal "GDZSChainsCountNeed", .ByID("GDZSChainsCountWork").Result + .ByID("GDZSChainsRezCountNeed").Result
                            .SetVal "GDZSMansCountNeed", .ByID("GDZSMansCountWork").Result + .ByID("GDZSMansRezCountNeed").Result
                        End If
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "В резерве" Then
                            .Raise "GDZSChainsRezCountHave"
                            .Raise "GDZSMansRezCountHave", CellVal(shp, "Prop.Personnel")
                        End If
                    Case Is = indexPers.ipGDZSPost 'ПБ
                        .Raise "GDZSPBCount"
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
                    Case Is = indexPers.ipKPP 'КПП
                        .Raise "GDZSKPPCount"
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
                    Case Is = indexPers.ipDimosos 'Дымосос
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
        
                '---Управление силами и средствами--------
                    Case Is = indexPers.ipBU 'БУ и СПР
                        If CellVal(shp, "Prop.UTP_STP_Reserv", visUnitsString) = "Участок" _
                            Then .Raise "BUCount"
                        If CellVal(shp, "Prop.UTP_STP_Reserv", visUnitsString) = "Сектор" _
                            Then .Raise "BUCount"
                    Case Is = indexPers.ipStab 'Штаб/РТП
                        .Raise "ShtabCount;PersonnelNeed"
                        If Not CellVal(shp, "Prop.NachStab", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.NachTyl", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.NachLabor", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.RTP2", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.RTP3", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.RTP4", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not CellVal(shp, "Prop.RTP5", visUnitsString) = "" _
                            Then .Raise "PersonnelNeed" ' pi_PersonnelNeed = pi_PersonnelNeed + 1
                    Case Is = indexPers.ipRNBD 'РНД
                        .Raise "RNBDCount"
        
                '---ПТВ-----------------------------------
                    Case Is = indexPers.ipStvolRuch 'Ручной водяной
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave"
                            
                            If shp.Cells("User.DiameterIn") = 50 Then
                                .Raise "StvolWBHave"
                            ElseIf shp.Cells("User.DiameterIn") = 70 Then
                                .Raise "StvolWAHave"
                            End If
                        End If
                    Case Is = indexPers.ipStvolRuchPena 'Ручной пенный
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .Raise "StvolFoamHave"
                            .Raise "FactStreamW", CellVal(shp, "Prop.PodOut") * (100 - CellVal(shp, "Prop.FoamPerc")) / 100
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipStvolGas 'Ручной газовый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolGasHave"
                        End If
                    Case Is = indexPers.ipStvolPoroshok 'Ручной порошковый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolPowderHave"
                        End If
                    Case Is = indexPers.ipStvolLafVoda 'Лафетный водяной
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave;StvolWLHave"
                        End If
                    Case Is = indexPers.ipStvolLafPena 'Лафетный пенный
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .Raise "StvolFoamHave"
                            .Raise "FactStreamW", CellVal(shp, "Prop.PodOut") * (100 - CellVal(shp, "Prop.FoamPerc")) / 100
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipStvolLafPoroshok 'Лафетный порошковый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolPowderHave"
                        End If
                    Case Is = indexPers.ipStvolLafVozimiy 'Лафетный водяной возимый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave;StvolWLHave"
                        End If
                    Case Is = indexPers.ipHydroelevator 'Гидроэлеватор
                        .RaiseByCellID shp, "Prop.Personnel"
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "User.PodOut")
                            .Raise "GetedWaterValueMax", CellVal(shp, "User.PodOut")
                            If IsEternalSource(shp) And .Result("WaterEternal") = False Then .SetVal "WaterEternal", True
                        End If
                        
                    Case Is = indexPers.ipPenosmesitelPerenosn 'Пеносмеситель
                        
                    Case Is = indexPers.ipPennayaVstavka 'Пенная вставка
                        .RaiseByCellID shp, "Prop.Personnel"
                    Case Is = indexPers.ipRazvetvlenie 'Разветвление
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestnPalka 'Лестница-палка ВОЗМОЖНО номер "85"???
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestnSturm 'Лестница-штурмовка
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestn3Kol 'Лестница-трехколенная
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipPenopodemnik 'Пеноподъемник
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipGrebenkaPGenerat 'Гребенка
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipKolonka 'Колонка
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "Prop.FlowMax")
                            .Raise "GetedWaterValueMax", CellVal(shp, "Prop.FlowMax")
                            .SetVal "WaterEternal", True
                        End If
                        
                    Case Is = indexPers.ipVsasLineWithSetk 'Всасывающая линия с сеткой
                        .RaiseByCellID shp, "Prop.Personnel"
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "Prop.PodIn")
                            .Raise "GetedWaterValueMax", CellVal(shp, "Prop.PodIn")
                            If IsEternalSource(shp) And .Result("WaterEternal") = False Then .SetVal "WaterEternal", True
                        End If
                        
                        
        
                '---Водоснабжение
                    Case Is = indexPers.ipPG 'ПГ
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipPW 'ПВ
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipPK 'ПК
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipOtkritiyVodoistochnik 'Водоем
                        .Raise "WaterSourceCount"
                        .Raise "WaterValueHave", CellVal(shp, "Prop.Value") * 1000
                    Case Is = indexPers.ipPirs 'Пирс
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipBashna 'Башня
                        .Raise "WaterSourceCount"
                        .Raise "WaterValueHave", CellVal(shp, "Prop.WTValue") * 1000
                    Case Is = indexPers.ipDistance 'Расстояние
                        .Raise "DistanceCount"
                        
                '---Линии
                    Case Is = indexPers.ipRukavLineNapor 'Напорная линия
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                                    And Not CellVal(shp, "Prop.LineType", visUnitsString) = "Рукав(4м)" Then
                            .Raise "LinesCount"   'pi_linesCount = pi_linesCount + 1
                            If CellVal(shp, "Prop.LineType", visUnitsString) = "Рабочая" Then
                                .Raise "WorklinesCount"
                                'Проверяем имеет ли данная фигура подпись положения
                                IsHoseShapeHavePos shp
                            ElseIf CellVal(shp, "Prop.LineType", visUnitsString) = "Магистральная" Then
                                .Raise "MaglinesCount"
                            End If
                            
                            tmp = CellVal(shp, "Prop.HoseDiameter", visUnitsString)
                            If shp.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "User.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "User.HosesNeed")
                            Else
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "Prop.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "Prop.HosesNeed")
                            End If
                            
                            If shp.CellExists("User.TotalLenight", 0) = True Then 'Имеет ли фигура ячейку "User.TotalLenight"
                                .Raise "HosesLength", CellVal(shp, "User.TotalLenight")
                            Else
                                .Raise "HosesLength", CellVal(shp, "User.LineLenightHose")
                            End If
                            'Проверяем имеет ли данная фигура подпись описания
                            IsHoseShapeHaveLabels shp

                        End If
                    Case Is = indexPers.ipSkatka 'Рукавная скатка
                        .Raise "HoseSkatka"
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                            And Not CellVal(shp, "Prop.LineType", visUnitsString) = "Рукав(4м)" Then

                            tmp = CellVal(shp, "Prop.HoseDiameter", visUnitsString)
                            If shp.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "User.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "User.HosesNeed")
                            Else
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "Prop.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "Prop.HosesNeed")
                            End If
                        End If
                    Case Is = indexPers.ipGarmoshka 'Рукавная гармошка
                        .Raise "HoseSkatka"
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                            And Not shp.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                            tmp = CellVal(shp, "Prop.HoseDiameter", visUnitsString)
                            If shp.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "User.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "User.HosesNeed")
                            Else
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "Prop.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "Prop.HosesNeed")
                            End If
                        End If
                    Case Is = indexPers.ipRukavLineVsas 'Всасывающая линия или напорно-всасывающая
                    Case Is = indexPers.ipPodpisPolozhenia 'Подпись положения линии
                        .Raise "LinesPosCount"
                    Case Is = indexPers.ipPodpisRukava 'Подпись рукавной линии
                        .Raise "LinesLableCount"
                    Case Is = indexPers.ipZdanie 'Здания
                        .Raise "BuildCount"
                    Case Is = indexPers.ipPodpisUniversal 'Подпись универсальная (так же подпись улиц)
                    Case Is = indexPers.ipRosaVetrov 'Roza vetrov
                        .Raise "OrientCount"
                        
                    Case Is = indexPers.ipKolichestvo 'Index kolichestva izdeliy techniky
                        .Raise "PersonnelHave", (CellVal(shp, "Prop.PersonnelHave") - CellVal(shp, "Prop.PersonnelHave") / CellVal(shp, "Prop.Quantity")) - (CellVal(shp, "Prop.Quantity") - 1)
                        .Raise "WaterValueHave", (CellVal(shp, "Prop.Water") - CellVal(shp, "Prop.Water") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses38Have", (CellVal(shp, "Prop.Hoses38") - CellVal(shp, "Prop.Hoses38") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses51Have", (CellVal(shp, "Prop.Hoses51") - CellVal(shp, "Prop.Hoses51") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses66Have", (CellVal(shp, "Prop.Hoses66") - CellVal(shp, "Prop.Hoses66") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses77Have", (CellVal(shp, "Prop.Hoses77") - CellVal(shp, "Prop.Hoses77") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses89Have", (CellVal(shp, "Prop.Hoses89") - CellVal(shp, "Prop.Hoses89") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses110Have", (CellVal(shp, "Prop.Hoses110") - CellVal(shp, "Prop.Hoses110") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses150Have", (CellVal(shp, "Prop.Hoses150") - CellVal(shp, "Prop.Hoses150") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses200Have", (CellVal(shp, "Prop.Hoses200") - CellVal(shp, "Prop.Hoses200") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses250Have", (CellVal(shp, "Prop.Hoses250") - CellVal(shp, "Prop.Hoses250") / CellVal(shp, "Prop.Quantity"))
                        .Raise "Hoses300Have", (CellVal(shp, "Prop.Hoses300") - CellVal(shp, "Prop.Hoses300") / CellVal(shp, "Prop.Quantity"))
                        
                        
'Количество
                        Select Case shp.Cells("User.MotherIndexPers")
                                  '---Пожарные автомобили-----------
                            Case Is = indexPers.ipAC 'Автоцистерны
                                .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipANR 'АНР
                                .Raise "MainPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAPP 'АПП
                                .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAL 'АЛ
                                .Raise "SpecialPAHave;ALCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKP 'АКП
                                .Raise "SpecialPAHave;AKPCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASO 'АСО
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAT 'АТ
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipADU 'АДУ
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPNS 'ПНС
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAA 'AA
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAV 'АВ
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKT 'АКТ
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAP 'АП
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAGVT 'АГВТ
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAG 'АГТ
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAGDZS 'АГДЗС
                                .Raise "TargetedPAHave;AGCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPK 'ПКС
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPozhLab 'ЛБ
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASA 'АСА
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASH 'АШ
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAR 'АР
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                                
                            Case Is = indexPers.ipACL 'АЦЛ
                                .Raise "MainPAHave;ALCount;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKP 'АЦКП
                                .Raise "MainPAHave;AKPCount;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAPP 'АПП
                                .Raise "MainPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                        
                        '---Прочая пожарная техника----------
                            Case Is = indexPers.ipPoezd 'Поезда
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipSamolet 'Самолет
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipVertolet 'Вертолет
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMotoPump 'Мотопомпы
                                .Raise "OtherTechincsHave;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipKorabl 'Корабль
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipKater 'Катер
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipTank 'Танки
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMVD 'МВД
                                .Raise "MVDCount;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
                                .Raise "SMPCount;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAvtoObschee 'Общее обозначение
                                If CellVal(shp, "Prop.RSCHS", visUnitsString) = "Входит" _
                                  Then .Raise "TotalTechCount", CellVal(shp, "Prop.Quantity") - 1 ' pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
                                .Raise "TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                                 On Error Resume Next
                                If CellVal(shp, "Prop.Dest", visUnitsString) = "Специальный" _
                                    Then .Raise "SpecialPAHave", CellVal(shp, "Prop.Quantity") - 1 ' pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, CellVal(shp, "Prop.Dest", visUnitsString), "общ") <> 0 _
                                    Then .Raise "MainPAHave", CellVal(shp, "Prop.Quantity") - 1 ' pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, CellVal(shp, "Prop.Dest", visUnitsString), "целев") <> 0 _
                                    Then .Raise "TargetedPAHave", CellVal(shp, "Prop.Quantity") - 1 ' pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, CellVal(shp, "Prop.Dest", visUnitsString), "Основ") <> 0 _
                                    Then .Raise "MainOverallHave", CellVal(shp, "Prop.Quantity") - 1 ' pi_MainOverallCount = pi_MainOverallCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                On Error GoTo EX
                        End Select
        
                    
                End Select
                'Вычисляем некоторые элементы по CellName
                .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77;Prop.Hose89;Prop.Hose110;Prop.Hose110;Prop.Hose150;Prop.Hose200;Prop.Hose250;Prop.Hose300"
                End With    'elements
                
'            End If
        End If
    Next shp


    
'Завершающий код - обобщает все полученные данные и по ним делат окончательный рассчет
'---Например кол-во людей необходимых для прокладки линий
    With elements
        'Суммируем имеющиеся рукава
        .SetVal "HosesHave", .Sum("Hoses38Have;Hoses51Have;Hoses66Have;Hoses77Have;Hoses89Have;Hoses110Have;Hoses150Have;Hoses200Have;Hoses250Have;Hoses300Have")
        'Прочие
        .SetVal "PAHighHave", .Sum("ALCount;AKPCount")
        .SetVal "TechnicsNotMchsHave", .Result("TotalTechCount") - .Result("TotalFireCount")
        .SetVal "TechnicsNotMchsOtherHave", .Result("TechnicsNotMchsHave") - .Result("MVDCount") - .Result("SMPCount")
        .SetVal "ACNeed", PF_RoundUp(.Result("PersonnelNeed") / 4)
        .SetVal "ANRNeed", PF_RoundUp(.Result("PersonnelNeed") / 5)
        .SetVal "StvolWBNeed", PF_RoundUp(.Result("NeedStreamW") / 3.7)
        .SetVal "StvolWANeed", PF_RoundUp(.Result("NeedStreamW") / 7.4)
        .SetVal "StvolWLNeed", PF_RoundUp(.Result("NeedStreamW") / 12)
        .SetVal "PANeedOnWaterSource", PF_RoundUp(.Result("NeedStreamW") / 32)
        .SetVal "RSCHSTechincsHave", .Result("TotalTechCount")
        .SetVal "WaterValueNeed10min", .Result("NeedStreamW") * 600
        
        'ГДЗС
        .SetVal "GDZSChainsCountWork", .Sum("GDZSChainsCountWorkSoft;GDZSChainsCountWorkHard")
        '---расчет требуемого кол-ва резервных ввеньев для работы в простых условиях
        If options.GDZSRezRoundUp Then
            .SetVal "GDZSChainsRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWorkSoft").Result / 3)
            .SetVal "GDZSMansRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWorkSoft").Result / 3) * 3
        Else
            .SetVal "GDZSChainsRezCountNeed", Int(.ByID("GDZSChainsCountWorkSoft").Result / 3)
            .SetVal "GDZSMansRezCountNeed", Int(.ByID("GDZSChainsCountWorkSoft").Result / 3) * 3
        End If
        '---расчет требуемого кол-ва резервных ввеньев для работы в сложных условиях
        .Raise "GDZSChainsRezCountNeed", .Result("GDZSChainsCountWorkHard")
        .Raise "GDZSMansRezCountNeed", .Result("GDZSChainsCountWorkHard") * 5
        '---Общее количесатво требуемых звеноьев
        .SetVal "GDZSChainsCountNeed", .Sum("GDZSChainsCountWork;GDZSChainsRezCountNeed")
        
        
        'Итог
        .Raise "PersonnelNeed", .Result("GDZSMansRezCountNeed")
    End With
    
    

'---Устанавливаем разницу значений начала процедуры и конца
'Конец функции
RaiseEvent AfterRefresh
Set Refresh = Me
Exit Function
EX:
    SaveLog Err, "Refresh"
    If Not shp Is Nothing Then Debug.Print shp.Name
    Set Refresh = Me
End Function



Private Function fb_FirstChain(as_Number As Single) As Boolean
'Функция возвращает Истина, если звено первое, и Ложь, если нет
    If as_Number - Fix(as_Number) < 0.4 And as_Number - Fix(as_Number) > 0 Then
        fb_FirstChain = True
    Else
        fb_FirstChain = False
    End If
End Function

Private Function IsHoseShapeHaveLabels(ByRef hoseShape As Visio.Shape) As Boolean
'Проверяем наличие у рукавной линии подписи положения и подписи
Dim con As Connect
Dim haveLabel As Boolean

    haveLabel = False

    For Each con In hoseShape.FromConnects
        If IsGFSShapeWithIP(con.FromSheet, 150, True) Then
            haveLabel = True
        End If
    Next con

    If elements.Result("AllHosesWithLabels") Then elements.SetVal "AllHosesWithLabels", haveLabel
    IsHoseShapeHaveLabels = haveLabel
End Function

Private Function IsHoseShapeHavePos(ByRef hoseShape As Visio.Shape) As Boolean
'Проверяем наличие у рукавной линии подписи положения и подписи
Dim con As Connect
Dim havePos As Boolean

    havePos = False

    For Each con In hoseShape.FromConnects
        If IsGFSShapeWithIP(con.FromSheet, 151, True) Then
            havePos = True
        End If
    Next con

    If elements.Result("AllHosesWithPos") Then elements.SetVal "AllHosesWithPos", havePos
    IsHoseShapeHavePos = havePos
End Function

Private Function IsEternalSource(ByRef shp As Visio.Shape) As Boolean
'Проверяем, звбирает ли воду гидроэлеватор или всасывающая сетка из открытого водоема с бесконечным запасом
Dim shpWS As Visio.Shape
Dim shpWSIndex As Long
    
    On Error GoTo EX
    
    shpWSIndex = CellVal(shp, "User.WSShapeID")
    If shpWSIndex > 0 Then
        Set shpWS = Application.ActivePage.Shapes.ItemFromID(shpWSIndex)
        If IsGFSShapeWithIP(shpWS, indexPers.ipOtkritiyVodoistochnik) Then
            IsEternalSource = CellVal(shpWS, "Prop.Type", visUnitsString) = "Неограниченный запас"
        End If
    End If
    
Exit Function
EX:
    IsEternalSource = False
End Function

'---------------Функции работы с элементами-----------------------
Public Function Result(ByVal ID As String) As Variant
'Получаем значение элемента с указанным ID
    Result = elements.Result(ID)
End Function
Public Function ResultByCN(ByVal callName As String) As Variant
'Получаем значение элемента с указанным ID
    ResultByCN = elements.ResultByCN(callName)
End Function

Public Function Sum(ByVal elemIDs As String) As Variant
'Получаем сумму значений элементов с указанными через точку с запятой именами элементов - elemIDs
    Sum = elements.Sum(elemIDs)
End Function

Public Sub PrintState(Optional ByVal elemIDs As String = "")
'Печатаем в Debug состояние указанных или всех элементов
    elements.PrintState elemIDs
End Sub
