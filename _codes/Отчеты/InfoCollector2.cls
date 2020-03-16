VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoCollector2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------Класс сбора, обобщения и обобщения информации имеющейся на листе------------------------------------------
Option Explicit


'----------------------------------------Коллекция элементов---------------------------------------------------------
Private elements As ElementsShell
'----------------------------------------Публичные свойства (настройки)----------------------------------------------
'Public targetPageIndex As Integer            'Целевой лист подсчета (0-все листы, 1-... - лист с указанным номером)
Public GDZSRezRoundUp As Boolean             'Округление резервных звеньев в большую сторону (True - для  каждого из блоков по три звена выставляется резерв, False - резерв выставляется при наличии очередных трех звеньев)


'----------------------------------------Внутренине свойства (служебные)---------------------------------------------
Private latestSquareTime As Date




'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
'    targetPageIndex = ActivePage.Index
    Set elements = New ElementsShell

'    VarInit
'    sC_Refresh (pi_TargetPageIndex)

End Sub
Private Sub Class_Terminate()
    Set elements = Nothing
End Sub








'Private Sub VarInit()
''Инициируются переменные
'    pb_AllHosesWithLabels = True
'    pb_AllHosesWithPos = True
'End Sub


Public Sub Refresh(a_PageIndex As Integer)
'Процедура обновления свойств класса
'Перебирает все элементы находящиеся на указанном a_PageIndex листе
Dim shp As Visio.Shape
Dim shpIndex As indexPers
Dim targetPage As Visio.Page
'Dim vs_Dest As String
'Dim tempTime1, tempTime2 As Date

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
'        If vsO_Shape.CellExists("User.IndexPers", 0) = True And vsO_Shape.CellExists("User.Version", 0) = True Then 'Является ли фигура фигурой ГраФиС
        If IsGFSShape(shp) Then
'            If vsO_Shape.Cells("User.Version") >= CP_GrafisVersion Then  'Проверяем версию фигуры
                shpIndex = shp.Cells("User.IndexPers")   'Определяем индекс фигуры ГраФиС
                
                With elements
                Select Case shpIndex
                '---Пожарные автомобили-----------
                    Case Is = ipAC   'Автоцистерны
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_MainPAHave = pi_MainPAHave + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_ACCount = pi_ACCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                        
                        
                        .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
                        .RaiseByCellID shp, "Prop.Hose38,Prop.Hose51,Prop.Hose66,Prop.Hose77"
                        
                        
                    Case Is = indexPers.ipANR 'АНР
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_MainPAHave = pi_MainPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipAKP 'АПП
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_MainPAHave = pi_MainPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipAL 'АЛ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_ALCount = pi_ALCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
    '                    pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipAKP 'АКП
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_AKPCount = pi_AKPCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
    '                    pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipASO 'АСО
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipAT 'АТ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipADU 'АДУ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipPNS 'ПНС
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                    Case Is = indexPers.ipAA 'AA
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipAV 'АВ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
                    Case Is = indexPers.ipAKT 'АКТ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipAP 'АП
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipAGVT 'АГВТ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipAG 'АГ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_TargetedPAHave = pi_TargetedPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipAGDZS 'АГДЗС
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_AGCount = pi_AGCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipPNS 'ПКС
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipPozhLab 'ЛБ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipASA 'АСА
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipASH 'АШ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipAR 'АР
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_SpecialPAHave = pi_SpecialPAHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                    Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                        '>>>Учет возможности отсутствия ячеек "Prop.Dest" для старых фигур
                        vs_Dest = CellVal(vsO_Shape, "Prop.Dest", visUnitsString)
                        If vs_Dest = "Специальный" _
                            Then pi_SpecialPAHave = pi_SpecialPAHave + 1
                        If InStr(1, vs_Dest, "общ") <> 0 _
                            Then pi_MainPAHave = pi_MainPAHave + 1
                        If InStr(1, vs_Dest, "целев") <> 0 _
                            Then pi_TargetedPAHave = pi_TargetedPAHave + 1
                        If InStr(1, vs_Dest, "Основ") <> 0 _
                            Then pi_MainOverallCount = pi_MainOverallCount + 1
    '                    On Error Resume Next
    '                    If vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString) = "Специальный" _
    '                        Then pi_SpecialPAHave = pi_SpecialPAHave + 1
    '                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "общ") <> 0 _
    '                        Then pi_MainPAHave = pi_MainPAHave + 1
    '                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "целев") <> 0 _
    '                        Then pi_TargetedPAHave = pi_TargetedPAHave + 1
    '                    If InStr(1, vsO_Shape.Cells("Prop.Dest").ResultStr(visUnitsString), "Основ") <> 0 _
    '                        Then pi_MainOverallCount = pi_MainOverallCount + 1
    '                    On Error GoTo EX
                        '<<<Учет возможности отсутствия ячейки "Prop.Dest" для старых фигур
                    Case Is = indexPers.ipAvtoObschee 'Общее обозначение
                        If vsO_Shape.Cells("Actions.RSCHS.Checked").Result(visNumber) = 1 _
                            Then pi_TechTotalCount = pi_TechTotalCount + 1
                        
    
                        
                    Case Is = indexPers.ipACL 'АЦЛ
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_MainPAHave = pi_MainPAHave + 1
                        pi_ALCount = pi_ALCount + 1
                        pi_ACCount = pi_ACCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipACKP 'АЦКП
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_MainPAHave = pi_MainPAHave + 1
                        pi_AKPCount = pi_AKPCount + 1
                        pi_ACCount = pi_ACCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
    
                
                '---Прочая пожарная техника----------
                    Case Is = indexPers.ipPoezd 'Поезда
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                    Case Is = indexPers.ipSamolet 'Самолет
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipVertolet 'Вертолет
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipMotoPump 'Мотопомпы
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                    Case Is = indexPers.ipKorabl 'Корабль
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipKater 'Катер
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipTank 'Танки
                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
                        'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        pi_FireTotalCount = pi_FireTotalCount + 1
                    Case Is = indexPers.ipMVD 'МВД
                        pi_MVDCount = pi_MVDCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                    Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
                        pi_MZdravCount = pi_MZdravCount + 1
                        pi_TechTotalCount = pi_TechTotalCount + 1
                        
                '---Объекты пожара-----------------
                    Case Is = indexPers.ipPloschadPozhara 'Площадь пожара
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
                        
                    Case Is = indexPers.ipZadiml 'Задымление
                        pi_SmokeCount = pi_SmokeCount + 1
                    Case Is = indexPers.ipOchag 'Очаг
                        pi_OchagCount = pi_OchagCount + 1
                    Case Is = indexPers.ipPostradavshie 'Гибель
                    Case Is = indexPers.ipRasprostran 'Распространение
                        pi_DevelopCount = pi_DevelopCount + 1
                    Case Is = indexPers.ipStorm 'Огненный шторм
                    Case Is = indexPers.ipOdinPozhar 'Одиночный пожар
                    Case Is = indexPers.ipPodpisStepenOgnestoykosty 'stepen ognestoikosti
                        pi_SOCount = pi_SOCount + 1
                
                '---СВязь и освещение------------------
                    Case Is = indexPers.ipProzhektor 'Прожектор
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    Case Is = indexPers.ipNosimoeRadio 'Радиостанция носимая
                    Case Is = indexPers.ipAvtoRadio 'Радиостанция возимая
                    Case Is = indexPers.ipStacRadio 'Радиостанция стационарная
                    Case Is = indexPers.ipGromkogovoritel 'Громкоговоритель
                    Case Is = indexPers.ipTelefon 'Телефон
        
                '---ГДЗС----------------------------
                    Case Is = indexPers.ipGDZSZvenoVozd 'Звено - ДАСВ
                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "Работает" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_GDZSChainsCount = pi_GDZSChainsCount + 1
                            pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
                            If vsO_Shape.Cells("Prop.WorkPlace").resultstr(visUnitsString) = "Стандартные условия" Then
                                If bo_GDZSRezRoundUp Then
                                    ps_GDZSChainsRezNeed = PF_RoundUp(pi_GDZSChainsCount / 3)
                                Else
                                    ps_GDZSChainsRezNeed = Fix(pi_GDZSChainsCount / 3)
                                End If
                            Else
                                ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
                                If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
                            End If
                        End If
                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "В резерве" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
                            pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
                        End If
                      
                        
                    Case Is = indexPers.ipGDZSZvenoKisl 'Звено - ДАСК
                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "Работает" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_GDZSChainsCount = pi_GDZSChainsCount + 1
                            pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
                            If vsO_Shape.Cells("Prop.WorkPlace").resultstr(visUnitsString) = "Стандартные условия" Then
                                If bo_GDZSRezRoundUp Then
                                    ps_GDZSChainsRezNeed = PF_RoundUp(pi_GDZSChainsCount / 3)
                                Else
                                    ps_GDZSChainsRezNeed = Fix(pi_GDZSChainsCount / 3)
                                End If
                            Else
                                ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
                                If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
                            End If
                        End If
                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "В резерве" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
                            pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
                        End If
                        
                    Case Is = indexPers.ipGDZSPost 'ПБ
                        pi_GDZSpbCount = pi_GDZSpbCount + 1
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    Case Is = indexPers.ipKPP 'КПП
                        pi_KPPCount = pi_KPPCount + 1
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    Case Is = indexPers.ipDimosos 'Дымосос
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
        
                '---Управление силами и средствами--------
                    Case Is = indexPers.ipBU 'БУ и СПР
                        If vsO_Shape.Cells("Prop.UTP_STP_Reserv").resultstr(visUnitsString) = "Участок" _
                            Then pi_BUCount = pi_BUCount + 1
                        If vsO_Shape.Cells("Prop.UTP_STP_Reserv").resultstr(visUnitsString) = "Сектор" _
                            Then pi_SPRCount = pi_SPRCount + 1
                    Case Is = indexPers.ipStab 'Штаб/РТП
                        pi_ShtabCount = pi_ShtabCount + 1
                        pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.NachStab").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.NachTyl").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.NachLabor").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.RTP2").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.RTP3").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.RTP4").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                        If Not vsO_Shape.Cells("Prop.RTP5").resultstr(visUnitsString) = "" _
                            Then pi_PersonnelNeed = pi_PersonnelNeed + 1
                    Case Is = indexPers.ipRNBD 'РНД
                        pi_RNBDCount = pi_RNBDCount + 1
        
                '---ПТВ-----------------------------------
                    Case Is = indexPers.ipStvolRuch 'Ручной водяной
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                            pi_StvolWHave = pi_StvolWHave + 1
                            If vsO_Shape.Cells("User.DiameterIn") = 50 Then
                                pi_StvolWBHave = pi_StvolWBHave + 1
                            ElseIf vsO_Shape.Cells("User.DiameterIn") = 70 Then
                                pi_StvolWAHave = pi_StvolWAHave + 1
                            End If
                        End If
                    Case Is = indexPers.ipStvolRuchPena 'Ручной пенный
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_StvolFoamHave = pi_StvolFoamHave + 1
                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                        End If
                    Case Is = indexPers.ipStvolGas 'Ручной газовый
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_StvolGasHave = pi_StvolGasHave + 1
                        End If
                    Case Is = indexPers.ipStvolPoroshok 'Ручной порошковый
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_StvolPowderHave = pi_StvolPowderHave + 1
                        End If
                    Case Is = indexPers.ipStvolLafVoda 'Лафетный водяной
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                            pi_StvolWHave = pi_StvolWHave + 1
                            pi_StvolWLHave = pi_StvolWLHave + 1
                        End If
                    Case Is = indexPers.ipStvolLafPena 'Лафетный пенный
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_StvolFoamHave = pi_StvolFoamHave + 1
                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                        End If
                    Case Is = indexPers.ipStvolLafPoroshok 'Лафетный порошковый
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            pi_StvolPowderHave = pi_StvolPowderHave + 1
                        End If
                    Case Is = indexPers.ipStvolLafVozimiy 'Лафетный водяной возимый
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
                            pi_StvolWHave = pi_StvolWHave + 1
                            pi_StvolWLHave = pi_StvolWLHave + 1
                        End If
                    Case Is = indexPers.ipHydroelevator 'Гидроэлеватор
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        If vsO_Shape.Cells("Prop.GetingWater").resultstr(visUnitsString) = "Да" Then
                            pi_GetingWaterCount = pi_GetingWaterCount + 1
                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodOut")
                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodOut")
                        End If
                        
                    Case Is = indexPers.ipPenosmesitelPerenosn 'Пеносмеситель
                        
                    Case Is = indexPers.ipPennayaVstavka 'Пенная вставка
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                    Case Is = indexPers.ipRazvetvlenie 'Разветвление
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        
                    Case Is = indexPers.ipLestnPalka 'Лестница-палка ВОЗМОЖНО номер "85"???
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        
                    Case Is = indexPers.ipLestnSturm 'Лестница-штурмовка
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        
                    Case Is = indexPers.ipLestn3Kol 'Лестница-трехколенная
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        
                    Case Is = indexPers.ipPenopodemnik 'Пеноподъемник
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        End If
                    Case Is = indexPers.ipGrebenkaPGenerat 'Гребенка
                        If vsO_Shape.Cells("Prop.MainManeure").resultstr(visUnitsString) = "Основное положение" Then
                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        End If
                    Case Is = indexPers.ipKolonka 'Колонка
                        If vsO_Shape.Cells("Prop.GetingWater").resultstr(visUnitsString) = "Да" Then
                            pi_GetingWaterCount = pi_GetingWaterCount + 1
                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.FlowMax")
                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.FlowMax")
                        End If
                        
                    Case Is = indexPers.ipVsasLineWithSetk 'Всасывающая линия с сеткой
                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        If vsO_Shape.Cells("Prop.GetingWater").resultstr(visUnitsString) = "Да" Then
                            pi_GetingWaterCount = pi_GetingWaterCount + 1
                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodIn")
                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodIn")
                        End If
                        
                        
        
                '---Водоснабжение
                    Case Is = indexPers.ipPG 'ПГ
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                    Case Is = indexPers.ipPV 'ПВ
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                    Case Is = indexPers.ipPK 'ПК
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                    Case Is = indexPers.ipOtkritiyVodoistochnik 'Водоем
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Value") * 1000
                    Case Is = indexPers.ipPirs 'Пирс
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                    Case Is = indexPers.ipBashna 'Башня
                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.WTValue") * 1000
                    Case Is = indexPers.ipDistance 'Расстояние
                        pi_distanceCount = pi_distanceCount + 1
                        
                '---Линии
                    Case Is = indexPers.ipRukavLineNapor 'Напорная линия
                        If vsO_Shape.Cells("Prop.ManeverHose").resultstr(visUnitsString) = "Нет" _
                            And Not vsO_Shape.Cells("Prop.LineType").resultstr(visUnitsString) = "Рукав(4м)" Then
                              pi_linesCount = pi_linesCount + 1
                            If vsO_Shape.Cells("Prop.LineType").resultstr(visUnitsString) = "Рабочая" Then
                                pi_WorklinesCount = pi_WorklinesCount + 1
                                'Проверяем имеет ли данная фигура подпись положения
    '                            Stop    'Реализовать!!!
    '                            IsHoseShapeHavePos vsO_Shape
                            End If
                            'If vsO_Shape.CellExists("User.Skatka", 0) = True Then pi_HoseSkatka = pi_HoseSkatka + 1 'считаем сколько рукавов в скатках
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
                                Case Is = 200
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 250
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 300
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
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
                    Case Is = indexPers.ipSkatka 'Рукавная скатка
                        pi_HoseSkatka = pi_HoseSkatka + 1
                        If vsO_Shape.Cells("Prop.ManeverHose").resultstr(visUnitsString) = "Нет" _
                            And Not vsO_Shape.Cells("Prop.LineType").resultstr(visUnitsString) = "Рукав(4м)" Then
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
                                Case Is = 200
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 250
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 300
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                            End Select
                        End If
                    Case Is = indexPers.ipGarmoshka 'Рукавная гармошка
                        pi_HoseSkatka = pi_HoseSkatka + 1
                        If vsO_Shape.Cells("Prop.ManeverHose").resultstr(visUnitsString) = "Нет" _
                            And Not vsO_Shape.Cells("Prop.LineType").resultstr(visUnitsString) = "Рукав(4м)" Then
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
                                Case Is = 200
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 250
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                                Case Is = 300
                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
                                    Else
                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
                                    End If
                            End Select
                        End If
                    Case Is = indexPers.ipRukavLineVsas 'Всасывающая линия или напорно-всасывающая
                    Case Is = indexPers.ipPodpisPolozhenia 'Подпись положения линии
                        pi_linesPosCount = pi_linesPosCount + 1
                    Case Is = indexPers.ipPodpisRukava 'Подпись рукавной линии
                        pi_linesLableCount = pi_linesLableCount + 1
                    Case Is = indexPers.ipZdanie 'Здания
                        pi_BuildCount = pi_BuildCount + 1
                    Case Is = indexPers.ipPodpisUniversal 'Подпись универсальная (так же подпись улиц)
                        pi_OrientCount = pi_OrientCount + 1
                    Case Is = indexPers.ipRosaVetrov 'Roza vetrov
                        pi_OrientCount = pi_OrientCount + 1
                        
                    Case Is = indexPers.ipKolichestvo 'Index kolichestva izdeliy techniky
                        pi_PersonnelHave = pi_PersonnelHave + (vsO_Shape.Cells("Prop.PersonnelHave") - vsO_Shape.Cells("Prop.PersonnelHave") / vsO_Shape.Cells("Prop.Quantity")) - (vsO_Shape.Cells("Prop.Quantity") - 1)
                        pi_WaterValueHave = pi_WaterValueHave + (vsO_Shape.Cells("Prop.Water") - vsO_Shape.Cells("Prop.Water") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses38Have = pi_Hoses38Have + (vsO_Shape.Cells("Prop.Hoses38") - vsO_Shape.Cells("Prop.Hoses38") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses66Have = pi_Hoses66Have + (vsO_Shape.Cells("Prop.Hoses66") - vsO_Shape.Cells("Prop.Hoses66") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses51Have = pi_Hoses51Have + (vsO_Shape.Cells("Prop.Hoses51") - vsO_Shape.Cells("Prop.Hoses51") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses77Have = pi_Hoses77Have + (vsO_Shape.Cells("Prop.Hoses77") - vsO_Shape.Cells("Prop.Hoses77") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses89Have = pi_Hoses89Have + (vsO_Shape.Cells("Prop.Hoses89") - vsO_Shape.Cells("Prop.Hoses89") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses110Have = pi_Hoses110Have + (vsO_Shape.Cells("Prop.Hoses110") - vsO_Shape.Cells("Prop.Hoses110") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses150Have = pi_Hoses150Have + (vsO_Shape.Cells("Prop.Hoses150") - vsO_Shape.Cells("Prop.Hoses150") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses200Have = pi_Hoses200Have + (vsO_Shape.Cells("Prop.Hoses200") - vsO_Shape.Cells("Prop.Hoses200") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses250Have = pi_Hoses250Have + (vsO_Shape.Cells("Prop.Hoses250") - vsO_Shape.Cells("Prop.Hoses250") / vsO_Shape.Cells("Prop.Quantity"))
                        pi_Hoses300Have = pi_Hoses300Have + (vsO_Shape.Cells("Prop.Hoses300") - vsO_Shape.Cells("Prop.Hoses300") / vsO_Shape.Cells("Prop.Quantity"))
                        
'Количество
                        Select Case vsO_Shape.Cells("User.MotherIndexPers")
                                  '---Пожарные автомобили-----------
                            Case Is = indexPers.ipAC 'Автоцистерны
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipANR 'АНР
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAPP 'АПП
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAL 'АЛ
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAKP 'АКП
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipASO 'АСО
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAT 'АТ
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipADU 'АДУ
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPNS 'ПНС
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAA 'AA
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAV 'АВ
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAKT 'АКТ
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAP 'АП
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAGVT 'АГВТ
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAG 'АГТ
                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAGDZS 'АГДЗС
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_AGCount = pi_AGCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPK 'ПКС
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPozhLab 'ЛБ
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipASA 'АСА
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipASH 'АШ
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAR 'АР
                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
            
                                
                            Case Is = indexPers.ipACL 'АЦЛ
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAKP 'АЦКП
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAPP 'АПП
                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                        
                        '---Прочая пожарная техника----------
                            Case Is = indexPers.ipPoezd 'Поезда
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipSamolet 'Самолет
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipVertolet 'Вертолет
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipMotoPump 'Мотопомпы
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipKorabl 'Корабль
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipKater 'Катер
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipTank 'Танки
                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipMVD 'МВД
                                pi_MVDCount = pi_MVDCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
                                pi_MZdravCount = pi_MZdravCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipAvtoObschee 'Общее обозначение
                                If vsO_Shape.Cells("Prop.RSCHS").resultstr(visUnitsString) = "Входит" _
                                  Then pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                 On Error Resume Next
                                If vsO_Shape.Cells("Prop.Dest").resultstr(visUnitsString) = "Специальный" _
                                    Then pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, vsO_Shape.Cells("Prop.Dest").resultstr(visUnitsString), "общ") <> 0 _
                                    Then pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, vsO_Shape.Cells("Prop.Dest").resultstr(visUnitsString), "целев") <> 0 _
                                    Then pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                If InStr(1, vsO_Shape.Cells("Prop.Dest").resultstr(visUnitsString), "Основ") <> 0 _
                                    Then pi_MainOverallCount = pi_MainOverallCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                On Error GoTo EX
                        End Select
        
                    
                End Select
                End With    'elements
                
'            End If
        End If
    Next vsO_Shape

   'Суммируем имеющиеся рукава
    pi_HosesHave = _
    pi_Hoses38Have + _
    pi_Hoses66Have + _
    pi_Hoses51Have + _
    pi_Hoses77Have + _
    pi_Hoses89Have + _
    pi_Hoses110Have + _
    pi_Hoses150Have + _
    pi_Hoses200Have + _
    pi_Hoses250Have + _
    pi_Hoses300Have
    
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







'Private Sub sc_SetPropertiesNull()
''Процедура обнуления свойств класса
''---ОБнуляем базовые свойства
'    pdt_LatestSquareTime = 0      'Крайнее значение времени площади
'    ps_FireSquare = 0             'Площадь пожара
'    ps_ExtSquare = 0              'Площадь тушения
'    ps_FactStreemW = 0            'Фактический расход воды
'    ps_IntenseWater = 0           'ТРебуемая интенсивность поачи воды - в данный момент не используется
'    ps_NeedStreemW = 0            'Требуемый расход воды
'    pi_PersonnelHave = 0          'Имеется личного состава
'    pi_PersonnelNeed = 0          'Требуется личного состава
'    pi_StvolWHave = 0             'Водяных стволов имеется
'    pi_MainPAHave = 0             'Основных ПА общего применения имеется
'    pi_TargetedPAHave = 0         'Основных ПА целевого применения имеется
'    pi_SpecialPAHave = 0          'Специальных ПА имеется
'    pi_OtherTechincsHave = 0      'Прочей техники имеется
'    ps_GetedWaterValue = 0        'Забирается воды из внешних источников
'    ps_GetedWaterValueMax = 0     'Забирается воды из внешних источников Максимум
'    pi_GetingWaterCount = 0       'Устройств забирающих воду кол-во
'    pi_GDZSChainsCount = 0        'Количество звеньев ГДЗС
'    pi_GDZSChainsRezCount = 0     'Количество резервных звеньев ГДЗС
'    pi_GDZSMansCount = 0          'Количество газрдымозащитников
'    pi_GDZSMansRezCount = 0       'Количество газодымозащитников для резервных звеньев ГДЗС
'    pi_GDZSpbCount = 0            'Kolichestvo PB
'    pi_KPPCount = 0               'Kolichestvo KPP
'    pi_Hoses51Count = 0           'Количество рукавов диаметром 51мм
'    pi_Hoses66Count = 0           'Количество рукавов диаметром 66мм
'    pi_Hoses77Count = 0           'Количество рукавов диаметром 77мм
'    pi_Hoses89Count = 0           'Количество рукавов диаметром 89мм
'    pi_Hoses110Count = 0          'Количество рукавов диаметром 110мм
'    pi_Hoses150Count = 0          'Количество рукавов диаметром 150мм
'    pi_Hoses200Count = 0          'Количество рукавов диаметром 200мм
'    pi_Hoses250Count = 0          'Количество рукавов диаметром 250мм
'    pi_Hoses300Count = 0          'Количество рукавов диаметром 300мм
'    pi_HosesCount = 0             'Количество рукавов общее
'    pi_Hoses38Have = 0            'Количество рукавов диаметром 38мм имеется
'    pi_Hoses51Have = 0            'Количество рукавов диаметром 51мм имеется
'    pi_Hoses66Have = 0            'Количество рукавов диаметром 66мм имеется
'    pi_Hoses77Have = 0            'Количество рукавов диаметром 77мм имеется
'    pi_Hoses89Have = 0            'Количество рукавов диаметром 89мм имеется
'    pi_Hoses110Have = 0           'Количество рукавов диаметром 110мм имеется
'    pi_Hoses150Have = 0           'Количество рукавов диаметром 150мм имеется
'    pi_Hoses200Have = 0           'Количество рукавов диаметром 200мм имеется
'    pi_Hoses250Have = 0           'Количество рукавов диаметром 250мм имеется
'    pi_Hoses300Have = 0           'Количество рукавов диаметром 300мм имеется
'    pi_HoseSkatka = 0             'Количество рукавов в скатках
'    pi_HosesCount = 0             'Количество рукавов общее
'    pi_HosesLength = 0            'Общая длина рукавных линий
'    pi_WaterValueNeed = 0         'Требуемый запас воды
'    pi_WaterValueHave = 0         'Фактический запас воды
'    pi_StvolWBHave = 0            'Водяных стволов Б подано
'    pi_StvolWAHave = 0            'Водяных стволов А подано
'    pi_StvolWLHave = 0            'Водяных лафетных стволов подано
'    pi_StvolFoamHave = 0          'Пенных стволов подано
'    pi_StvolGasHave = 0           'Газовых стволов подано
'    pi_StvolPowderHave = 0        'Порошковых стволов подано
'    pi_WaterValueHave = 0         'Требуемый запас воды
'    pi_WaterValueHave = 0         'Фактический запас воды
'    pi_ACCount = 0                'АЦ
'    pi_AGCount = 0                'АГ
'    pi_ALCount = 0                'АЛ
'    pi_AKPCount = 0               'КП
'    pi_MVDCount = 0               'МВД
'    pi_MZdravCount = 0            'СМП
'    pi_TechTotalCount = 0         'Общее количество техники
'    pi_FireTotalCount = 0         'Пожарная
'    pi_BUCount = 0                'Боевых участков
'    pi_SPRCount = 0               'Секторов
'    pi_ShtabCount = 0             'Shtab
'    pi_RNBDCount = 0              'Reshaushee napravlenie
'    pi_OchagCount = 0             'Ochag
'    pi_DevelopCount = 0           'Rasprostranenie
'    pi_SmokeCount = 0             'Zadimlenie
'    pi_WaterSourceCount = 0       'Vodoistochniki
'    pi_distanceCount = 0          'Rasstoyanie
'    pi_linesCount = 0             'Otdelnie linii
'    pi_linesPosCount = 0          'Polozenie linii
'    pi_linesLableCount = 0        'Podpisi linii
'    pi_WorklinesCount = 0         'Rabochie linii
'    pi_BuildCount = 0             'Zdania
'    pi_SOCount = 0                'Stepen ognestoikosti
'    pi_FireCount = 0              'Ploshad gorenia
'    pi_OrientCount = 0            'Ориентиры
'    pb_GDZSDiscr = False          'Несоответствие человек сложным условиям
'    ps_GDZSChainsRezNeed = 0      'Требуется звеньев
'
'
'
'
''---Обнуляем расчетные свойства
''    pi_StvolWNeed = 0        'Водяных стволов требуется
''    pi_MainPANeed = 0      'Основных ПА общего применения требуется
''    pi_GetingWaterNeedCount = 0    'Устройств забирающих воду кол-во - требуется
''    pi_GDZSChainsNeedCount = 0  'Количество звеньев ГДЗС - требуется
''    pi_RezGDZSChainsNeedCount = 0  'Количество резервных звеньев ГДЗС - требуется
'
'End Sub

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
'        Debug.Print con.FromSheet.Name
    Next con
    
    If pb_AllHosesWithPos Then pb_AllHosesWithPos = tmpVal
End Sub



