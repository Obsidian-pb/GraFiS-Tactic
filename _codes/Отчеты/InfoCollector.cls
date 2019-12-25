VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------Класс сбора, обобщения и обобщения информации имеющейся на листе------------------------------------------
Option Explicit

'----------------------------------------Служебные свойства----------------------------------------------------------
Public pi_TargetPageIndex As Integer          'Целевой лист подсчета (0-все листы, 1-... - лист с указанным номером)
'Private WithEvents vco_App As Visio.Application

'----------------------------------------Целевые свойства------------------------------------------------------------
'----------------------------------------Базовые---------------------------------------------------------------------
Public pdt_LatestSquareTime As Date   'Время самой поздней площади пожара
Public ps_FireSquare As Single        'Площадь пожара
Public ps_ExtSquare As Single         'Площадь тушения
Public ps_FactStreemW As Single       'Фактический расход воды
Public ps_IntenseWater As Single      'ТРебуемая интенсивность поачи воды - в данный момент не используется
Public ps_NeedStreemW As Single       'Требуемый расход воды


Public pi_PersonnelHave As Integer    'Имеется личного состава
Public pi_PersonnelNeed As Integer    'Требуется личного состава

Public pi_StvolWHave As Integer       'Водяных стволов подано
Public pi_StvolWBHave As Integer      'Водяных стволов Б подано
Public pi_StvolWAHave As Integer      'Водяных стволов А подано
Public pi_StvolWLHave As Integer      'Водяных лафетных стволов подано

Public pi_StvolFoamHave As Integer    'Пенных стволов подано
Public pi_StvolGasHave As Integer     'Газовых стволов подано
Public pi_StvolPowderHave As Integer  'Порошковых стволов подано


Public pi_MainPAHave As Integer       'Основных ПА общего применения имеется
Public pi_TargetedPAHave As Integer   'Основных ПА целевого применения имеется
Public pi_MainOverallCount As Integer 'Основная
Public pi_SpecialPAHave As Integer    'Специальных ПА имеется
Public pi_OtherTechincsHave As Integer    'Прочей техники имеется


Public ps_GetedWaterValue As Single       'Забирается воды из внешних источников
Public ps_GetedWaterValueMax As Single    'Забирается воды из внешних источников Максимум
Public pi_GetingWaterCount As Integer   'Устройств забирающих воду кол-во


Public pi_GDZSChainsCount As Integer 'Количество рабочих звеньев ГДЗС
Public pi_GDZSMansCount As Integer 'Количество газрдымозащитников
Public pi_GDZSChainsRezCount As Integer 'Количество резервных звеньев ГДЗС
Public pi_GDZSMansRezCount As Integer 'Количество газодымозащитников для резервных звеньев ГДЗС
Public ps_GDZSChainsRezNeed As Single 'Требуемое количество звеньев
Public pb_GDZSDiscr As Boolean 'Несоотвествие количества газодымозащитников в звене сложным условиям
Public pi_GDZSpbCount As Integer 'Kolichestvo PB
Public pi_KPPCount As Integer 'Kolichestvo KPP


Public pi_Hoses51Count As Integer  'Количество рукавов диаметром 51мм
Public pi_Hoses66Count As Integer  'Количество рукавов диаметром 66мм
Public pi_Hoses77Count As Integer  'Количество рукавов диаметром 77мм
Public pi_Hoses89Count As Integer  'Количество рукавов диаметром 89мм
Public pi_Hoses110Count As Integer 'Количество рукавов диаметром 110мм
Public pi_Hoses150Count As Integer 'Количество рукавов диаметром 150мм
Public pi_HosesCount As Integer    'Количество всех напорных рукавов
Public pi_HosesLength As Integer   'Общая длина рукавных линий


Public pi_WaterValueNeed As Long   'Требуемый запас воды
Public pi_WaterValueHave As Long   'Фактический запас воды

'---Новые------------
Public pi_ACCount As Integer            'АЦ
Public pi_AGCount As Integer            'АГ
Public pi_ALCount As Integer            'АЛ
Public pi_AKPCount As Integer           'КП
Public pi_MVDCount As Integer           'МВД
Public pi_MZdravCount As Integer        'СМП
Public pi_TechTotalCount As Integer     'Общее количество техники
Public pi_FireTotalCount As Integer     'Пожарная
Public pi_BUCount As Integer            'Боевых участков
Public pi_ShtabCount As Integer         'Shtab
Public pi_RNBDCount As Integer          'Reshaushee napravlenie
Public pi_OchagCount As Integer         'Ochag
Public pi_DevelopCount As Integer       'Rasprostranenie
Public pi_SmokeCount As Integer         'Zadimlenie
Public pi_WaterSourceCount As Integer   'Vodoistochniki
Public pi_distanceCount As Integer      'Rasstoyanie
Public pi_linesCount As Integer         'Otdelnie linii
Public pi_linesPosCount As Integer      'Polozenie linii
Public pi_linesLableCount As Integer    'Podpisi linii
Public pi_WorklinesCount As Integer     'Rabochie linii
Public pi_BuildCount As Integer         'Zdania
Public pi_SOCount As Integer            'Stepen ognestoikosti
Public pi_FireCount As Integer          'Ploshad gorenia
Public pi_OrientCount As Integer        'Ориентиры

'---Новые2------------
Public pb_AllHosesWithLabels As Boolean         'Все рукавные линии с подписями
Public pb_AllHosesWithPos As Boolean            'Все рабочие рукавные линии сподписями положения








'----------------------------------------Расчетные----------------------------------------------------------------------
'Public pi_StvolWNeed As Integer       'Водяных стволов требуется
'Public pi_MainPANeed As Integer       'Основных ПА общего применения требуется
'Public pi_GetingWaterNeedCount As Integer   'Устройств забирающих воду кол-во - требуется
'Public pi_GDZSChainsNeedCount As Integer 'Количество звеньев ГДЗС - требуется
'Public pi_RezGDZSChainsNeedCount As Integer 'Количество резервных звеньев ГДЗС - требуется




'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
    pi_TargetPageIndex = ActivePage.Index
    VarInit
    'Set vco_App = Visio.Application
    sC_Refresh (pi_TargetPageIndex)

End Sub


Private Sub Class_Terminate()
'Set vco_App = Nothing
'MsgBox "Класс уничтожен"
End Sub

Private Sub VarInit()
'Инициируются переменные
    pb_AllHosesWithLabels = True
    pb_AllHosesWithPos = True
End Sub


Public Sub sC_Refresh(as_PageIndex As Integer)
'Процедура обновления свойств класса
'Перебирает все элементы находящиеся на указанном as_PageIndex листе
Dim vsO_Shape As Visio.Shape
Dim vsi_ShapeIndex As Integer
Dim vsO_TargetPage As Visio.Page
Dim tempTime1, tempTime2 As Date

    On Error GoTo EX
'Обнуляем количественные свойства
'---Обнуляем минимальное время
pdt_LatestSquareTime = 0
'---Обнуляем прочие количественные свойства
sc_SetPropertiesNull

Set vsO_TargetPage = Application.ActiveDocument.Pages(as_PageIndex)
'MsgBox vsO_TargetPage.Name

'---Перебираем все фигуры и в случае если фигура является фигурой ГраФиС анализируем её
For Each vsO_Shape In vsO_TargetPage.Shapes
    If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then 'Является ли фигура фигурой ГраФиС
        If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then  'Проверяем версию фигуры
            vsi_ShapeIndex = vsO_Shape.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
            
            Select Case vsi_ShapeIndex
            '---Пожарные автомобили-----------
                Case Is = 1 'Автоцистерны
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_MainPAHave = pi_MainPAHave + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_ACCount = pi_ACCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 2 'АНР
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_MainPAHave = pi_MainPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 163 'АПП
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_MainPAHave = pi_MainPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                Case Is = 3 'АЛ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_ALCount = pi_ALCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 4 'АКП
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_AKPCount = pi_AKPCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 5 'АСО
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 6 'АТ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 7 'АДУ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 8 'ПНС
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 9 'AA
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 10 'АВ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 11 'АКТ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 12 'АП
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 13 'АГВТ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 14 'АГТ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_TargetedPAHave = pi_TargetedPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 15 'АГДЗС
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_AGCount = pi_AGCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 16 'ПКС
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 17 'ЛБ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 18 'АСА
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 19 'АШ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 20 'АР
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_SpecialPAHave = pi_SpecialPAHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 160 'Общее обозначение
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    '>>>Учет возможности отсутствия ячеек "Prop.Dest" для старых фигур
                    On Error Resume Next
                    If vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString) = "Специальный" _
                        Then pi_SpecialPAHave = pi_SpecialPAHave + 1
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "общ") <> 0 _
                        Then pi_MainPAHave = pi_MainPAHave + 1
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "целев") <> 0 _
                        Then pi_TargetedPAHave = pi_TargetedPAHave + 1
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "Основ") <> 0 _
                        Then pi_MainOverallCount = pi_MainOverallCount + 1
                    On Error GoTo EX
                    '<<<Учет возможности отсутствия ячейки "Prop.Dest" для старых фигур
                Case Is = 33 'Общее обозначение
                    If vsO_Shape.Cells("Actions.RSCHS.Checked").Result(visNumber) = 1 _
                        Then pi_TechTotalCount = pi_TechTotalCount + 1
                    

                    
                Case Is = 161 'АЦЛ
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_MainPAHave = pi_MainPAHave + 1
                    pi_ALCount = pi_ALCount + 1
                    pi_ACCount = pi_ACCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                Case Is = 162 'АЦКП
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_MainPAHave = pi_MainPAHave + 1
                    pi_AKPCount = pi_AKPCount + 1
                    pi_ACCount = pi_ACCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")

            
            '---Прочая пожарная техника----------
                Case Is = 24 'Поезда
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 25 'Самолет
                   'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 26 'Самолет-амфибия
                   'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 27 'Вертолет
                   'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 28 'Мотопомпы
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                Case Is = 30 'Корабль
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 31 'Катер
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 73 'Машины на гусеничном ходу
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 74 'Танки
                    pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                    pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                    'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    pi_FireTotalCount = pi_FireTotalCount + 1
                Case Is = 3001 'МВД
                    pi_MVDCount = pi_MVDCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                Case Is = 3002 'Минздрав
                    pi_MZdravCount = pi_MZdravCount + 1
                    pi_TechTotalCount = pi_TechTotalCount + 1
                    
            '---Объекты пожара-----------------
                Case Is = 64 'Площадь пожара
                      pi_FireCount = pi_FireCount + 1
                    If vsO_Shape.Cells("Prop.SquareTime") > pdt_LatestSquareTime Then 'Проверяем является ли данная фигура последней по времени
                        ps_FireSquare = vsO_Shape.Cells("User.FireSquare")
                        ps_ExtSquare = vsO_Shape.Cells("User.ExtSquare")
                        ps_NeedStreemW = vsO_Shape.Cells("User.NeedStreamW")
                                            
                        pdt_LatestSquareTime = vsO_Shape.Cells("Prop.SquareTime") 'Устанавливаем новое значение наибольшего времени учета для площади
                    ElseIf vsO_Shape.Cells("Prop.SquareTime") = pdt_LatestSquareTime Then
                        ps_FireSquare = ps_FireSquare + vsO_Shape.Cells("User.FireSquare")
                        ps_ExtSquare = ps_ExtSquare + vsO_Shape.Cells("User.ExtSquare")
                        ps_NeedStreemW = ps_NeedStreemW + vsO_Shape.Cells("User.NeedStreamW")
                                            
                    End If
                    
                Case Is = 71 'Задымление
                    pi_SmokeCount = pi_SmokeCount + 1
                Case Is = 70 'Очаг
                    pi_OchagCount = pi_OchagCount + 1
                Case Is = 92 'Гибель
                Case Is = 84 'Распространение
                    pi_DevelopCount = pi_DevelopCount + 1
                Case Is = 86 'Огненный шторм
                Case Is = 87 'Одиночный пожар
                Case Is = 153 'stepen ognestoikosti
                    pi_SOCount = pi_SOCount + 1
            
            '---СВязь и освещение------------------
                Case Is = 57 'Прожектор
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                Case Is = 58 'Радиостанция носимая
                Case Is = 59 'Радиостанция возимая
                Case Is = 23 'Радиостанция стационарная
                Case Is = 60 'Громкоговоритель
                Case Is = 61 'Телефон
    
            '---ГДЗС----------------------------
                Case Is = 46 'Звено - ДАСВ
                    If vsO_Shape.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_GDZSChainsCount = pi_GDZSChainsCount + 1
                        pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
                          If vsO_Shape.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                              ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 0.3334
                          Else
                              ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
                              If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
                          End If
                    End If
                    If vsO_Shape.Cells("Prop.WorkState").ResultStr(visUnitsString) = "В резерве" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
                        pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
                    End If
                  
                    
                Case Is = 90 'Звено - ДАСК
                    If vsO_Shape.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_GDZSChainsCount = pi_GDZSChainsCount + 1
                        pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
                          If vsO_Shape.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                              ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 0.3334
                          Else
                              ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
                              If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
                          End If
                    End If
                    If vsO_Shape.Cells("Prop.WorkState").ResultStr(visUnitsString) = "В резерве" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
                        pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
                    End If
                    
                Case Is = 47 'ПБ
                    pi_GDZSpbCount = pi_GDZSpbCount + 1
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                Case Is = 48 'КПП
                    pi_KPPCount = pi_KPPCount + 1
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                Case Is = 49 'Дымосос
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
    
            '---Управление силами и средствами--------
                Case Is = 63 'БУ
                    pi_BUCount = pi_BUCount + 1
                Case Is = 62 'Штаб/РТП
                    pi_ShtabCount = pi_ShtabCount + 1
                    pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.NachStab").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.NachTyl").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.NachLabor").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.RTP2").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.RTP3").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.RTP4").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    If Not vsO_Shape.Cells("Prop.RTP5").ResultStr(visUnitsString) = "" _
                        Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                Case Is = 83 'РНД
                    pi_RNBDCount = pi_RNBDCount + 1
    
            '---ПТВ-----------------------------------
                Case Is = 34 'Ручной водяной
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                        pi_StvolWHave = pi_StvolWHave + 1
                        If vsO_Shape.Cells("User.DiameterIn") = 50 Then
                            pi_StvolWBHave = pi_StvolWBHave + 1
                        ElseIf vsO_Shape.Cells("User.DiameterIn") = 70 Then
                            pi_StvolWAHave = pi_StvolWAHave + 1
                        End If
                    End If
                Case Is = 35 'Ручной пенный
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_StvolFoamHave = pi_StvolFoamHave + 1
                        ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                    End If
                Case Is = 76 'Ручной газовый
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_StvolGasHave = pi_StvolGasHave + 1
                    End If
                Case Is = 77 'Ручной порошковый
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_StvolPowderHave = pi_StvolPowderHave + 1
                    End If
                Case Is = 36 'Лафетный водяной
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                        pi_StvolWHave = pi_StvolWHave + 1
                        pi_StvolWLHave = pi_StvolWLHave + 1
                    End If
                Case Is = 37 'Лафетный пенный
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_StvolFoamHave = pi_StvolFoamHave + 1
                        ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                    End If
                Case Is = 38 'Лафетный порошковый
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_StvolPowderHave = pi_StvolPowderHave + 1
                    End If
                Case Is = 39 'Лафетный водяной возимый
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                        pi_StvolWHave = pi_StvolWHave + 1
                        pi_StvolWLHave = pi_StvolWLHave + 1
                    End If
                Case Is = 40 'Гидроэлеватор
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    If vsO_Shape.Cells("Prop.GetingWater").ResultStr(visUnitsString) = "Да" Then
                        pi_GetingWaterCount = pi_GetingWaterCount + 1
                        ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodOut")
                        ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodOut")
                    End If
                    
                Case Is = 41 'Пеносмеситель
                    
                Case Is = 191 'Пенная вставка
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                Case Is = 42 'Разветвление
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    
                Case Is = 21 'Лестница-палка ВОЗМОЖНО номер "85"???
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    
                Case Is = 43 'Лестница-штурмовка
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    
                Case Is = 44 'Лестница-трехколенная
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    
                Case Is = 45 'Пеноподъемник
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    End If
                Case Is = 22 'Гребенка
                    If vsO_Shape.Cells("Prop.MainManeure").ResultStr(visUnitsString) = "Основное положение" Then
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    End If
                Case Is = 72 'Колонка
                    If vsO_Shape.Cells("Prop.GetingWater").ResultStr(visUnitsString) = "Да" Then
                        pi_GetingWaterCount = pi_GetingWaterCount + 1
                        ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.FlowMax")
                        ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.FlowMax")
                    End If
                    
                Case Is = 88  'Всасывающая линия с сеткой
                    pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    If vsO_Shape.Cells("Prop.GetingWater").ResultStr(visUnitsString) = "Да" Then
                        pi_GetingWaterCount = pi_GetingWaterCount + 1
                        ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodIn")
                        ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodIn")
                    End If
                    
                    
    
            '---Водоснабжение
                Case Is = 50 'ПГ
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                Case Is = 51 'ПВ
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                Case Is = 52 'ПК
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                Case Is = 53 'Водоем
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Value") * 1000
                Case Is = 54 'Пирс
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                Case Is = 56 'Башня
                    pi_WaterSourceCount = pi_WaterSourceCount + 1
                    pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.WTValue") * 1000
                Case Is = 75 'Rasstoyanie
                    pi_distanceCount = pi_distanceCount + 1
                    
            '---Линии
                Case Is = 100 'Напорная линия
                    If vsO_Shape.Cells("Prop.ManeverHose").ResultStr(visUnitsString) = "Нет" _
                        And Not vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
                          pi_linesCount = pi_linesCount + 1
                                                If vsO_Shape.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рабочая" Then
                            pi_WorklinesCount = pi_WorklinesCount + 1
                            'Проверяем имеет ли данная фигура подпись положения
'                            Stop    'Реализовать!!!
                            IsHoseShapeHavePos vsO_Shape
                        End If
                        Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
                            Case Is = 51
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                            Case Is = 66
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                            Case Is = 77
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                            Case Is = 89
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                            Case Is = 110
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                            Case Is = 150
                                If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                    pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("User.HosesNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                Else
                                    pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("Prop.HoseNeed")
                                    pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                End If
                        End Select
                        If vsO_Shape.CellExists("User.TotalLenight", 0) = True Then 'Имеет ли фигура ячейку "User.TotalLenight"
                            pi_HosesLength = pi_HosesLength + vsO_Shape.Cells("User.TotalLenight")
                        Else
                            pi_HosesLength = pi_HosesLength + vsO_Shape.Cells("Prop.LineLenightHose")
                        End If
                        'Проверяем имеет ли данная фигура подпись описания
                        'Stop    'Реализовать!!!
                        pb_AllHosesWithLabels = True
                    End If
                Case Is = 101 'Всасывающая линия или напорно-всасывающая
                Case Is = 151 'Polozenie linii
                    pi_linesPosCount = pi_linesPosCount + 1
                Case Is = 150 'Podpis linii
                    pi_linesLableCount = pi_linesLableCount + 1
                Case Is = 135 'Zdania
                    pi_BuildCount = pi_BuildCount + 1
                Case Is = 152 'Podis ulici
                    pi_OrientCount = pi_OrientCount + 1
                Case Is = 91 'Roza vetrov
                    pi_OrientCount = pi_OrientCount + 1
                    
                Case Is = 195 'Index kolichestva izdeliy techniky
                    pi_PersonnelHave = pi_PersonnelHave + (vsO_Shape.Cells("Prop.PersonnelHave") - vsO_Shape.Cells("Prop.PersonnelHave") / vsO_Shape.Cells("Prop.Quantity")) - (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_WaterValueHave = pi_WaterValueHave + (vsO_Shape.Cells("Prop.Water") - vsO_Shape.Cells("Prop.Water") / vsO_Shape.Cells("Prop.Quantity"))
                    
                Select Case vsO_Shape.Cells("User.MotherIndexPers")
                          '---Пожарные автомобили-----------
                Case Is = 1 'Автоцистерны
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 2 'АНР
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 163 'АПП
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 3 'АЛ
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 4 'АКП
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 5 'АСО
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 6 'АТ
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 7 'АДУ
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 8 'ПНС
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 9 'AA
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 10 'АВ
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 11 'АКТ
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 12 'АП
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 13 'АГВТ
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 14 'АГТ
                    pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 15 'АГДЗС
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_AGCount = pi_AGCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 16 'ПКС
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 17 'ЛБ
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 18 'АСА
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 19 'АШ
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 20 'АР
                    pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)

                    
                Case Is = 161 'АЦЛ
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 162 'АЦКП
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 163 'АПП
                    pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
            
            '---Прочая пожарная техника----------
                Case Is = 24 'Поезда
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 25 'Самолет
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 26 'Самолет-амфибия
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 27 'Вертолет
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 28 'Мотопомпы
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 30 'Корабль
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 31 'Катер
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 73 'Машины на гусеничном ходу
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 74 'Танки
                    pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 3001 'МВД
                    pi_MVDCount = pi_MVDCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 3002 'Минздрав
                    pi_MZdravCount = pi_MZdravCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 33 'Общее обозначение
                    If vsO_Shape.Cells("Prop.RSCHS").ResultStr(visUnitsString) = "Входит" _
                      Then pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                Case Is = 160 'Общее обозначение
                    pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                     On Error Resume Next
                    If vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString) = "Специальный" _
                        Then pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "общ") <> 0 _
                        Then pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "целев") <> 0 _
                        Then pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "Основ") <> 0 _
                        Then pi_MainOverallCount = pi_MainOverallCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                    On Error GoTo EX
                    End Select
    
                
            End Select
            
        End If
    End If
Next vsO_Shape

'Завершающий код - обобщает все полученные данные и по ним делат окончательный рассчет
'---Например кол-во людей необходимых для прокладки линий



'---Устанавливаем разницу значений начала процедуры и конца
'tempTime2 = Time()
'MsgBox tempTime2 - tempTime1
'MsgBox "123"
'Debug.Print tempTime2 - tempTime1

Exit Sub
EX:
    SaveLog Err, "sC_Refresh"
    Debug.Print vsO_Shape.Name
End Sub







Private Sub sc_SetPropertiesNull()
'Процедура обнуления свойств класса
'---ОБнуляем базовые свойства
    pdt_LatestSquareTime = 0      'Крайнее значение времени площади
    ps_FireSquare = 0             'Площадь пожара
    ps_ExtSquare = 0              'Площадь тушения
    ps_FactStreemW = 0            'Фактический расход воды
    ps_IntenseWater = 0           'ТРебуемая интенсивность поачи воды - в данный момент не используется
    ps_NeedStreemW = 0            'Требуемый расход воды
    pi_PersonnelHave = 0          'Имеется личного состава
    pi_PersonnelNeed = 0          'Требуется личного состава
    pi_StvolWHave = 0             'Водяных стволов имеется
    pi_MainPAHave = 0             'Основных ПА общего применения имеется
    pi_TargetedPAHave = 0         'Основных ПА целевого применения имеется
    pi_SpecialPAHave = 0          'Специальных ПА имеется
    pi_OtherTechincsHave = 0      'Прочей техники имеется
    ps_GetedWaterValue = 0        'Забирается воды из внешних источников
    ps_GetedWaterValueMax = 0     'Забирается воды из внешних источников Максимум
    pi_GetingWaterCount = 0       'Устройств забирающих воду кол-во
    pi_GDZSChainsCount = 0        'Количество звеньев ГДЗС
    pi_GDZSChainsRezCount = 0     'Количество резервных звеньев ГДЗС
    pi_GDZSMansCount = 0          'Количество газрдымозащитников
    pi_GDZSMansRezCount = 0       'Количество газодымозащитников для резервных звеньев ГДЗС
    pi_GDZSpbCount = 0            'Kolichestvo PB
    pi_KPPCount = 0               'Kolichestvo KPP
    pi_Hoses51Count = 0           'Количество рукавов диаметром 51мм
    pi_Hoses66Count = 0           'Количество рукавов диаметром 66мм
    pi_Hoses77Count = 0           'Количество рукавов диаметром 77мм
    pi_Hoses89Count = 0           'Количество рукавов диаметром 89мм
    pi_Hoses110Count = 0          'Количество рукавов диаметром 110мм
    pi_Hoses150Count = 0          'Количество рукавов диаметром 150мм
    pi_HosesCount = 0             'Количество рукавов общее
    pi_HosesLength = 0            'Общая длина рукавных линий
    pi_WaterValueNeed = 0         'Требуемый запас воды
    pi_WaterValueHave = 0         'Фактический запас воды
    pi_StvolWBHave = 0            'Водяных стволов Б подано
    pi_StvolWAHave = 0            'Водяных стволов А подано
    pi_StvolWLHave = 0            'Водяных лафетных стволов подано
    pi_StvolFoamHave = 0          'Пенных стволов подано
    pi_StvolGasHave = 0           'Газовых стволов подано
    pi_StvolPowderHave = 0        'Порошковых стволов подано
    pi_WaterValueHave = 0         'Требуемый запас воды
    pi_WaterValueHave = 0         'Фактический запас воды
    pi_ACCount = 0                'АЦ
    pi_AGCount = 0                'АГ
    pi_ALCount = 0                'АЛ
    pi_AKPCount = 0               'КП
    pi_MVDCount = 0               'МВД
    pi_MZdravCount = 0            'СМП
    pi_TechTotalCount = 0         'Общее количество техники
    pi_FireTotalCount = 0         'Пожарная
    pi_BUCount = 0                'Боевых участков
    pi_ShtabCount = 0             'Shtab
    pi_RNBDCount = 0              'Reshaushee napravlenie
    pi_OchagCount = 0             'Ochag
    pi_DevelopCount = 0           'Rasprostranenie
    pi_SmokeCount = 0             'Zadimlenie
    pi_WaterSourceCount = 0       'Vodoistochniki
    pi_distanceCount = 0          'Rasstoyanie
    pi_linesCount = 0             'Otdelnie linii
    pi_linesPosCount = 0          'Polozenie linii
    pi_linesLableCount = 0        'Podpisi linii
    pi_WorklinesCount = 0         'Rabochie linii
    pi_BuildCount = 0             'Zdania
    pi_SOCount = 0                'Stepen ognestoikosti
    pi_FireCount = 0              'Ploshad gorenia
    pi_OrientCount = 0            'Ориентиры
    pb_GDZSDiscr = False          'Несоответствие человек сложным условиям
    ps_GDZSChainsRezNeed = 0      'Требуется звеньев
   
    
    
    
'---Обнуляем расчетные свойства
'    pi_StvolWNeed = 0        'Водяных стволов требуется
'    pi_MainPANeed = 0      'Основных ПА общего применения требуется
'    pi_GetingWaterNeedCount = 0    'Устройств забирающих воду кол-во - требуется
'    pi_GDZSChainsNeedCount = 0  'Количество звеньев ГДЗС - требуется
'    pi_RezGDZSChainsNeedCount = 0  'Количество резервных звеньев ГДЗС - требуется

End Sub

Private Function fb_FirstChain(as_Number As Single) As Boolean
'Функция возвращает Истина, если звено первое, и Ложь, если нет
    If as_Number - Fix(as_Number) < 0.4 And as_Number - Fix(as_Number) > 0 Then
        fb_FirstChain = True
    Else
        fb_FirstChain = False
    End If
End Function

Private Sub IsHoseShapeHavePos(ByRef hoseShape As Visio.Shape)
'Проверяем наличие у рукавной линии подписи положения
Dim con As Connect
Dim tmpVal As Boolean

    tmpVal = False
    
    For Each con In hoseShape.FromConnects
        If con.FromSheet.Cells("User.IndexPers").Result(visNumber) = 151 Then tmpVal = True
        Debug.Print con.FromSheet.Name
    Next con
    
    If pb_AllHosesWithPos Then pb_AllHosesWithPos = tmpVal
End Sub

