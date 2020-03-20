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
Public elements As ElementsShell
'----------------------------------------Публичные свойства (настройки)----------------------------------------------
Public options As InfoCollectorOptions          'Настройки расчета

'----------------------------------------Внутренине свойства (служебные)---------------------------------------------
'Private latestSquareTime As Date
'----------------------------------------События класса---------------------------------------------------------
Public Event AfterRefresh()


'----------------------------------------Служебные процедуры------------------------------------------------------------
Private Sub Class_Initialize()
    Set elements = New ElementsShell
    Set options = New InfoCollectorOptions
'    VarInit

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
'Dim vs_Dest As String
'Dim tempTime1, tempTime2 As Date
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
                        
                        .Raise "MainOverallHave;MainPAHave;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipANR 'АНР
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_MainPAHave = pi_MainPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;MainPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipAPP 'АПП
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_MainPAHave = pi_MainPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;MainPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipAL 'АЛ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_ALCount = pi_ALCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
'    '                    pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "SpecialPAHave;ALCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipAKP 'АКП
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_AKPCount = pi_AKPCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") 'vsO_Shape.Cells("Prop.Hose51")
'    '                    pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") 'vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") 'vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "SpecialPAHave;AKPCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipASO 'АСО
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAT 'АТ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipADU 'АДУ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPNS 'ПНС
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
'                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
'                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
'                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
'                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
'                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
'                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                        
                        .Raise "TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77;Prop.Hose89;Prop.Hose110;Prop.Hose110;Prop.Hose150;Prop.Hose200;Prop.Hose250;Prop.Hose300"
                    Case Is = indexPers.ipAA 'AA
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipAV 'АВ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
'                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
'                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
'                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77;Prop.Hose89;Prop.Hose110;Prop.Hose150"
                    Case Is = indexPers.ipAKT 'АКТ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipAP 'АП
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAGVT 'АГВТ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAG 'АГ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_TargetedPAHave = pi_TargetedPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "MainOverallHave;TargetedPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAGDZS 'АГДЗС
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_AGCount = pi_AGCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipPNS 'ПКС
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipPozhLab 'ЛБ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipASA 'АСА
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipASH 'АШ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "SpecialPAHave;AGCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipAR 'АР
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_SpecialPAHave = pi_SpecialPAHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
'                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
'                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
'                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
'                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
'                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
'                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                        
                        .Raise "SpecialPAHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77;Prop.Hose89;Prop.Hose110;Prop.Hose150;Prop.Hose200;Prop.Hose250;Prop.Hose300"
                    Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
'                        pi_Hoses89Have = pi_Hoses89Have + CellVal(vsO_Shape, "Prop.Hose89") ' vsO_Shape.Cells("Prop.Hose89")
'                        pi_Hoses110Have = pi_Hoses110Have + CellVal(vsO_Shape, "Prop.Hose110") ' vsO_Shape.Cells("Prop.Hose110")
'                        pi_Hoses150Have = pi_Hoses150Have + CellVal(vsO_Shape, "Prop.Hose150") ' vsO_Shape.Cells("Prop.Hose150")
'                        pi_Hoses200Have = pi_Hoses200Have + CellVal(vsO_Shape, "Prop.Hose200") ' vsO_Shape.Cells("Prop.Hose200")
'                        pi_Hoses250Have = pi_Hoses250Have + CellVal(vsO_Shape, "Prop.Hose250") ' vsO_Shape.Cells("Prop.Hose250")
'                        pi_Hoses300Have = pi_Hoses300Have + CellVal(vsO_Shape, "Prop.Hose300") ' vsO_Shape.Cells("Prop.Hose300")
                        
                        .Raise "TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77;Prop.Hose89;Prop.Hose110;Prop.Hose150;Prop.Hose200;Prop.Hose250;Prop.Hose300"
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
                        If CellVal(shp, "Actions.RSCHS.Checked") = 1 _
                            Then .Raise "TotalTechCount" ' pi_TechTotalCount = pi_TechTotalCount + 1
                        
    
                        
                    Case Is = indexPers.ipACL 'АЦЛ
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_MainPAHave = pi_MainPAHave + 1
'                        pi_ALCount = pi_ALCount + 1
'                        pi_ACCount = pi_ACCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;MainPAHave;ALCount;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipACKP 'АЦКП
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_MainPAHave = pi_MainPAHave + 1
'                        pi_AKPCount = pi_AKPCount + 1
'                        pi_ACCount = pi_ACCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_Hoses38Have = pi_Hoses38Have + CellVal(vsO_Shape, "Prop.Hose38") ' vsO_Shape.Cells("Prop.Hose38")
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "MainOverallHave;MainPAHave;AKPCount;ACCount;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose38;Prop.Hose51;Prop.Hose66;Prop.Hose77"
                
                
                '---Прочая пожарная техника----------
                    Case Is = indexPers.ipPoezd 'Поезда
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
'                        pi_Hoses51Have = pi_Hoses51Have + CellVal(vsO_Shape, "Prop.Hose51") ' vsO_Shape.Cells("Prop.Hose51")
'                        pi_Hoses66Have = pi_Hoses66Have + CellVal(vsO_Shape, "Prop.Hose66") ' vsO_Shape.Cells("Prop.Hose66")
'                        pi_Hoses77Have = pi_Hoses77Have + CellVal(vsO_Shape, "Prop.Hose77") ' vsO_Shape.Cells("Prop.Hose77")
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                        .Raise "WaterValueHave", shp.Cells("Prop.Water")
'                        .RaiseByCellID shp, "Prop.Hose51;Prop.Hose66;Prop.Hose77"
                    Case Is = indexPers.ipSamolet 'Самолет
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipVertolet 'Вертолет
                       'pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                    Case Is = indexPers.ipMotoPump 'Мотопомпы
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalFireCount"
                        .Raise "PersonnelNeed", shp.Cells("Prop.Personnel")
                    Case Is = indexPers.ipKorabl 'Корабль
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipKater 'Катер
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipTank 'Танки
'                        pi_PersonnelHave = pi_PersonnelHave + vsO_Shape.Cells("Prop.PersonnelHave") - 1
'                        pi_OtherTechincsHave = pi_OtherTechincsHave + 1
'                        'pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Water")
'                        pi_TechTotalCount = pi_TechTotalCount + 1
'                        pi_FireTotalCount = pi_FireTotalCount + 1
                        
                        .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount"
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipMVD 'МВД
'                        pi_MVDCount = pi_MVDCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
                        
                        .Raise "TotalTechCount;MVDCount"
                    Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
'                        pi_MZdravCount = pi_MZdravCount + 1
'                        pi_TechTotalCount = pi_TechTotalCount + 1
                        
                        .Raise "TotalTechCount;SMPCount"
                        
                '---Объекты пожара-----------------
                    Case Is = indexPers.ipPloschadPozhara 'Площадь пожара
'                          pi_FireCount = pi_FireCount + 1
                        .Raise "FireCount"
'                        If vsO_Shape.Cells("Prop.SquareTime") > pdt_LatestSquareTime Then 'Проверяем является ли данная фигура последней по времени
                        If CellVal(shp, "Prop.SquareTime", visDate) > latestSquareTime Then 'Проверяем является ли данная фигура последней по времени
'                            ps_FireSquare = vsO_Shape.Cells("User.FireSquare")
'                            ps_ExtSquare = vsO_Shape.Cells("User.ExtSquare")
'                            ps_NeedStreemW = vsO_Shape.Cells("User.NeedStreamW")
                            .SetValByCellID shp, "User.FireSquare;User.ExtSquare;User.NeedStreamW"
                                                
                            latestSquareTime = shp.Cells("Prop.SquareTime") 'Устанавливаем новое значение наибольшего времени учета для площади
                        ElseIf shp.Cells("Prop.SquareTime") = latestSquareTime Then
'                            ps_FireSquare = ps_FireSquare + vsO_Shape.Cells("User.FireSquare")
'                            ps_ExtSquare = ps_ExtSquare + vsO_Shape.Cells("User.ExtSquare")
'                            ps_NeedStreemW = ps_NeedStreemW + vsO_Shape.Cells("User.NeedStreamW")
                            
                            .RaiseByCellID shp, "User.FireSquare;User.ExtSquare;User.NeedStreamW"
                        End If
                        
                    Case Is = indexPers.ipZadiml 'Задымление
'                        pi_SmokeCount = pi_SmokeCount + 1
                        .Raise "SmokeCount"
                    Case Is = indexPers.ipOchag 'Очаг
'                        pi_OchagCount = pi_OchagCount + 1
                        .Raise "OchagCount"
                    Case Is = indexPers.ipPostradavshie 'Гибель
                    Case Is = indexPers.ipRasprostran 'Распространение
'                        pi_DevelopCount = pi_DevelopCount + 1
                        .Raise "SpreadCount"
                    Case Is = indexPers.ipStorm 'Огненный шторм
                    Case Is = indexPers.ipOdinPozhar 'Одиночный пожар
                    Case Is = indexPers.ipPodpisStepenOgnestoykosty 'stepen ognestoikosti
'                        pi_SOCount = pi_SOCount + 1
                        .Raise "SOCount"
                
                '---СВязь и освещение------------------
                    Case Is = indexPers.ipProzhektor 'Прожектор
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .Raise "PersonnelHave", shp.Cells("Prop.PersonnelHave") - 1
                    Case Is = indexPers.ipNosimoeRadio 'Радиостанция носимая
                    Case Is = indexPers.ipAvtoRadio 'Радиостанция возимая
                    Case Is = indexPers.ipStacRadio 'Радиостанция стационарная
                    Case Is = indexPers.ipGromkogovoritel 'Громкоговоритель
                    Case Is = indexPers.ipTelefon 'Телефон
        
                '---ГДЗС----------------------------
                    Case Is = indexPers.ipGDZSZvenoVozd 'Звено - ДАСВ
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_GDZSChainsCount = pi_GDZSChainsCount + 1
'                            pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
                            
                            .Raise "GDZSChainsCountWork"
                            .Raise "PersonnelNeed;GDZSMansCountWork", CellVal(shp, "Prop.Personnel")
                            
                            If shp.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                                If options.GDZSRezRoundUp Then
'                                    ps_GDZSChainsRezNeed = PF_RoundUp(pi_GDZSChainsCount / 3)
                                    .SetVal "GDZSChainsRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWork").Result / 3)
                                    .SetVal "GDZSMansRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWork").Result / 3) * 3
                                Else
'                                    ps_GDZSChainsRezNeed = Fix(pi_GDZSChainsCount / 3)
                                    
                                    .Raise "GDZSChainsRezCountNeed", Int(.ByID("GDZSChainsCountWork").Result / 3)
                                    .Raise "GDZSMansRezCountNeed", Int(.ByID("GDZSChainsCountWork").Result / 3) * 3
                                End If
                            Else
'                                ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
                                .Raise "GDZSChainsRezCountNeed"
                                
'                                If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
                                If shp.Cells("Prop.Personnel") < 5 Then .SetVal "GDZSDiscr", True
                            End If
                            'Вычисляем требуемое количество звеньев и людей
                            .SetVal "GDZSChainsCountNeed", .ByID("GDZSChainsCountWork").Result + .ByID("GDZSChainsRezCountNeed").Result
                            .SetVal "GDZSMansCountNeed", .ByID("GDZSMansCountWork").Result + .ByID("GDZSMansRezCountNeed").Result
                        End If
                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "В резерве" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
'                            pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
                            
                            'ПРОВЕРИТЬ!!! Получается, что дважды считается требуемое количество людей
                            .Raise "GDZSMansRezCountHave", CellVal(shp, "Prop.Personnel")
                            .Raise "GDZSChainsRezCountHave"
                        End If
                        
                        
                    Case Is = indexPers.ipGDZSZvenoKisl 'Звено - ДАСК
'                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "Работает" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_GDZSChainsCount = pi_GDZSChainsCount + 1
'                            pi_GDZSMansCount = pi_GDZSMansCount + vsO_Shape.Cells("Prop.Personnel")
'                            If vsO_Shape.Cells("Prop.WorkPlace").resultstr(visUnitsString) = "Стандартные условия" Then
'                                If bo_GDZSRezRoundUp Then
'                                    ps_GDZSChainsRezNeed = PF_RoundUp(pi_GDZSChainsCount / 3)
'                                Else
'                                    ps_GDZSChainsRezNeed = Fix(pi_GDZSChainsCount / 3)
'                                End If
'                            Else
'                                ps_GDZSChainsRezNeed = ps_GDZSChainsRezNeed + 1
'                                If vsO_Shape.Cells("Prop.Personnel") < 5 Then pb_GDZSDiscr = True
'                            End If
'                        End If
'                        If vsO_Shape.Cells("Prop.WorkState").resultstr(visUnitsString) = "В резерве" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_GDZSChainsRezCount = pi_GDZSChainsRezCount + 1
'                            pi_GDZSMansRezCount = pi_GDZSMansRezCount + vsO_Shape.Cells("Prop.Personnel")
'                        End If

                        If shp.Cells("Prop.WorkState").ResultStr(visUnitsString) = "Работает" Then
                            .Raise "GDZSChainsCountWork"
                            .Raise "PersonnelNeed;GDZSMansCountWork", CellVal(shp, "Prop.Personnel")
                            If shp.Cells("Prop.WorkPlace").ResultStr(visUnitsString) = "Стандартные условия" Then
                                If options.GDZSRezRoundUp Then
                                    .Raise "GDZSChainsRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWork").Result / 3)
                                    .Raise "GDZSMansRezCountNeed", PF_RoundUp(.ByID("GDZSChainsCountWork").Result / 3) * 3
                                Else
                                    .Raise "GDZSChainsRezCountNeed", Int(.ByID("GDZSChainsCountWork").Result / 3)
                                    .Raise "GDZSMansRezCountNeed", Int(.ByID("GDZSChainsCountWork").Result / 3) * 3
                                End If
                            Else
                                .Raise "GDZSChainsRezCountNeed"
                                If shp.Cells("Prop.Personnel") < 5 Then .SetVal "GDZSDiscr", True
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
'                        pi_GDZSpbCount = pi_GDZSpbCount + 1
'                        pi_PersonnelNeed = pi_PersonnelNeed + CellVal(shp, "Prop.Personnel")
                        .Raise "GDZSPBCount"
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
                    Case Is = indexPers.ipKPP 'КПП
'                        pi_KPPCount = pi_KPPCount + 1
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .Raise "GDZSKPPCount"
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
                    Case Is = indexPers.ipDimosos 'Дымосос
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .Raise "PersonnelNeed", CellVal(shp, "Prop.Personnel")
        
                '---Управление силами и средствами--------
                    Case Is = indexPers.ipBU 'БУ и СПР
'                        If vsO_Shape.Cells("Prop.UTP_STP_Reserv").resultstr(visUnitsString) = "Участок" _
'                            Then pi_BUCount = pi_BUCount + 1
'                        If vsO_Shape.Cells("Prop.UTP_STP_Reserv").resultstr(visUnitsString) = "Сектор" _
'                            Then pi_SPRCount = pi_SPRCount + 1
                        If CellVal(shp, "Prop.UTP_STP_Reserv", visUnitsString) = "Участок" _
                            Then .Raise "BUCount"
                        If CellVal(shp, "Prop.UTP_STP_Reserv", visUnitsString) = "Сектор" _
                            Then .Raise "BUCount"
                    Case Is = indexPers.ipStab 'Штаб/РТП
'                        pi_ShtabCount = pi_ShtabCount + 1
'                        pi_PersonnelNeed = pi_PersonnelNeed + 1
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
'                        pi_RNBDCount = pi_RNBDCount + 1
                        .Raise "RNBDCount"
        
                '---ПТВ-----------------------------------
                    Case Is = indexPers.ipStvolRuch 'Ручной водяной
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
'                            pi_StvolWHave = pi_StvolWHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave"
                            
                            If shp.Cells("User.DiameterIn") = 50 Then
'                                pi_StvolWBHave = pi_StvolWBHave + 1
                                .Raise "StvolWBHave"
                            ElseIf shp.Cells("User.DiameterIn") = 70 Then
'                                pi_StvolWAHave = pi_StvolWAHave + 1
                                .Raise "StvolWAHave"
                            End If
                        End If
                    Case Is = indexPers.ipStvolRuchPena 'Ручной пенный
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_StvolFoamHave = pi_StvolFoamHave + 1
'                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                            
                            .Raise "StvolFoamHave"
                            .Raise "FactStreamW", CellVal(shp, "Prop.PodOut") * (100 - CellVal(shp, "Prop.FoamPerc")) / 100
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipStvolGas 'Ручной газовый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_StvolGasHave = pi_StvolGasHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolGasHave"
                        End If
                    Case Is = indexPers.ipStvolPoroshok 'Ручной порошковый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_StvolPowderHave = pi_StvolPowderHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolPowderHave"
                        End If
                    Case Is = indexPers.ipStvolLafVoda 'Лафетный водяной
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
'                            pi_StvolWHave = pi_StvolWHave + 1
'                            pi_StvolWLHave = pi_StvolWLHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave;StvolWLHave"
                        End If
                    Case Is = indexPers.ipStvolLafPena 'Лафетный пенный
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_StvolFoamHave = pi_StvolFoamHave + 1
'                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("Prop.PodOut") * (100 - vsO_Shape.Cells("Prop.FoamPerc")) / 100
                            
                            .Raise "StvolFoamHave"
                            .Raise "FactStreamW", CellVal(shp, "Prop.PodOut") * (100 - CellVal(shp, "Prop.FoamPerc")) / 100
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipStvolLafPoroshok 'Лафетный порошковый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            pi_StvolPowderHave = pi_StvolPowderHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel"
                            .Raise "StvolPowderHave"
                        End If
                    Case Is = indexPers.ipStvolLafVozimiy 'Лафетный водяной возимый
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
'                            ps_FactStreemW = ps_FactStreemW + vsO_Shape.Cells("User.PodOut")
'                            pi_StvolWHave = pi_StvolWHave + 1
'                            pi_StvolWLHave = pi_StvolWLHave + 1
                            
                            .RaiseByCellID shp, "Prop.Personnel;User.PodOut"
                            .Raise "StvolWHave;StvolWLHave"
                        End If
                    Case Is = indexPers.ipHydroelevator 'Гидроэлеватор
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
'                            pi_GetingWaterCount = pi_GetingWaterCount + 1
'                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodOut")
'                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodOut")
                            
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "User.PodOut")
                            .Raise "GetedWaterValueMax", CellVal(shp, "User.PodOut")
                        End If
                        
                    Case Is = indexPers.ipPenosmesitelPerenosn 'Пеносмеситель
                        
                    Case Is = indexPers.ipPennayaVstavka 'Пенная вставка
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                    Case Is = indexPers.ipRazvetvlenie 'Разветвление
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestnPalka 'Лестница-палка ВОЗМОЖНО номер "85"???
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestnSturm 'Лестница-штурмовка
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipLestn3Kol 'Лестница-трехколенная
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        
                    Case Is = indexPers.ipPenopodemnik 'Пеноподъемник
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipGrebenkaPGenerat 'Гребенка
                        If CellVal(shp, "Prop.MainManeure", visUnitsString) = "Основное положение" Then
'                            pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                            .RaiseByCellID shp, "Prop.Personnel"
                        End If
                    Case Is = indexPers.ipKolonka 'Колонка
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
'                            pi_GetingWaterCount = pi_GetingWaterCount + 1
'                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.FlowMax")
'                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.FlowMax")
                            
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "Prop.FlowMax")
                            .Raise "GetedWaterValueMax", CellVal(shp, "Prop.FlowMax")
                        End If
                        
                    Case Is = indexPers.ipVsasLineWithSetk 'Всасывающая линия с сеткой
'                        pi_PersonnelNeed = pi_PersonnelNeed + vsO_Shape.Cells("Prop.Personnel")
                        .RaiseByCellID shp, "Prop.Personnel"
                        If CellVal(shp, "Prop.GetingWater", visUnitsString) = "Да" Then
'                            pi_GetingWaterCount = pi_GetingWaterCount + 1
'                            ps_GetedWaterValue = ps_GetedWaterValue + vsO_Shape.Cells("Prop.PodIn")
'                            ps_GetedWaterValueMax = ps_GetedWaterValueMax + vsO_Shape.Cells("Prop.PodIn")
                            
                            .Raise "GetingWaterCount"
                            .Raise "GetedWaterValue", CellVal(shp, "Prop.PodIn")
                            .Raise "GetedWaterValueMax", CellVal(shp, "Prop.PodIn")
                        End If
                        
                        
        
                '---Водоснабжение
                    Case Is = indexPers.ipPG 'ПГ
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipPW 'ПВ
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipPK 'ПК
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipOtkritiyVodoistochnik 'Водоем
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.Value") * 1000
                        
                        .Raise "WaterSourceCount"
                        .Raise "WaterValueHave", CellVal(shp, "Prop.Value") * 1000
                    Case Is = indexPers.ipPirs 'Пирс
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
                        .Raise "WaterSourceCount"
                    Case Is = indexPers.ipBashna 'Башня
'                        pi_WaterSourceCount = pi_WaterSourceCount + 1
'                        pi_WaterValueHave = pi_WaterValueHave + vsO_Shape.Cells("Prop.WTValue") * 1000
                        
                        .Raise "WaterSourceCount"
                        .Raise "WaterValueHave", CellVal(shp, "Prop.WTValue") * 1000
                    Case Is = indexPers.ipDistance 'Расстояние
'                        pi_distanceCount = pi_distanceCount + 1
                        .Raise "DistanceCount"
                        
                '---Линии
                    Case Is = indexPers.ipRukavLineNapor 'Напорная линия
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                                    And Not CellVal(shp, "Prop.LineType", visUnitsString) = "Рукав(4м)" Then
                            .Raise "LinesCount"   'pi_linesCount = pi_linesCount + 1
                            If CellVal(shp, "Prop.LineType", visUnitsString) = "Рабочая" Then
'                                pi_WorklinesCount = pi_WorklinesCount + 1
                                .Raise "WorklinesCount"
                                'Проверяем имеет ли данная фигура подпись положения
                                IsHoseShapeHavePos shp
                            ElseIf CellVal(shp, "Prop.LineType", visUnitsString) = "Магистральная" Then
                                .Raise "MaglinesCount"
                            End If
                            'If vsO_Shape.CellExists("User.Skatka", 0) = True Then pi_HoseSkatka = pi_HoseSkatka + 1 'считаем сколько рукавов в скатках
'                            Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
'                                Case Is = 51
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 66
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 77
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 89
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 110
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 150
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 200
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 250
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 300
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                            End Select
                            
                            tmp = CellVal(shp, "Prop.HoseDiameter", visUnitsString)
                            If shp.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "User.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "User.HosesNeed")
                            Else
                                .Raise "Hoses" & tmp & "Count", CellVal(shp, "Prop.HosesNeed")
                                .Raise "HosesCount", CellVal(shp, "Prop.HosesNeed")
                            End If
                            
                            If shp.CellExists("User.TotalLenight", 0) = True Then 'Имеет ли фигура ячейку "User.TotalLenight"
'                                pi_HosesLength = pi_HosesLength + vsO_Shape.Cells("User.TotalLenight")
                                .Raise "HosesLength", CellVal(shp, "User.TotalLenight")
                            Else
'                                pi_HosesLength = pi_HosesLength + vsO_Shape.Cells("Prop.LineLenightHose")
                                .Raise "HosesLength", CellVal(shp, "User.LineLenightHose")
                            End If
                            'Проверяем имеет ли данная фигура подпись описания
                            IsHoseShapeHaveLabels shp

'                            .SetVal "AllHosesWithLabels", True
                        End If
                    Case Is = indexPers.ipSkatka 'Рукавная скатка
'                        pi_HoseSkatka = pi_HoseSkatka + 1
                        .Raise "HoseSkatka"
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                            And Not CellVal(shp, "Prop.LineType", visUnitsString) = "Рукав(4м)" Then
'                            Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
'                                Case Is = 51
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 66
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 77
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 89
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 110
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 150
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 200
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 250
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 300
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                            End Select

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
'                        pi_HoseSkatka = pi_HoseSkatka + 1
                        .Raise "HoseSkatka"
                        If CellVal(shp, "Prop.ManeverHose", visUnitsString) = "Нет" _
                            And Not shp.Cells("Prop.LineType").ResultStr(visUnitsString) = "Рукав(4м)" Then
'                            Select Case vsO_Shape.Cells("Prop.HoseDiameter")   '.ResultStr(visUnitsString)
'                                Case Is = 51
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses51Count = pi_Hoses51Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 66
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses66Count = pi_Hoses66Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 77
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses77Count = pi_Hoses77Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 89
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses89Count = pi_Hoses89Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 110
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses110Count = pi_Hoses110Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 150
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses150Count = pi_Hoses150Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 200
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses200Count = pi_Hoses200Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 250
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses250Count = pi_Hoses250Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                                Case Is = 300
'                                    If vsO_Shape.CellExists("User.HosesNeed", 0) = True Then 'Имеет ли фигура ячейку "User.HosesNeed"
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("User.HosesNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("User.HosesNeed")
'                                    Else
'                                        pi_Hoses300Count = pi_Hoses300Count + vsO_Shape.Cells("Prop.HoseNeed")
'                                        pi_HosesCount = pi_HosesCount + vsO_Shape.Cells("Prop.HoseNeed")
'                                    End If
'                            End Select
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
'                        pi_linesPosCount = pi_linesPosCount + 1
                        .Raise "LinesPosCount"
                    Case Is = indexPers.ipPodpisRukava 'Подпись рукавной линии
'                        pi_linesLableCount = pi_linesLableCount + 1
                        .Raise "LinesLableCount"
                    Case Is = indexPers.ipZdanie 'Здания
'                        pi_BuildCount = pi_BuildCount + 1
                        .Raise "BuildCount"
                    Case Is = indexPers.ipPodpisUniversal 'Подпись универсальная (так же подпись улиц)
'                        pi_OrientCount = pi_OrientCount + 1
                    Case Is = indexPers.ipRosaVetrov 'Roza vetrov
'                        pi_OrientCount = pi_OrientCount + 1
                        .Raise "OrientCount"
                        
                    Case Is = indexPers.ipKolichestvo 'Index kolichestva izdeliy techniky
'                        pi_PersonnelHave = pi_PersonnelHave + (vsO_Shape.Cells("Prop.PersonnelHave") - vsO_Shape.Cells("Prop.PersonnelHave") / vsO_Shape.Cells("Prop.Quantity")) - (vsO_Shape.Cells("Prop.Quantity") - 1)
'                        pi_WaterValueHave = pi_WaterValueHave + (vsO_Shape.Cells("Prop.Water") - vsO_Shape.Cells("Prop.Water") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses38Have = pi_Hoses38Have + (vsO_Shape.Cells("Prop.Hoses38") - vsO_Shape.Cells("Prop.Hoses38") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses66Have = pi_Hoses66Have + (vsO_Shape.Cells("Prop.Hoses66") - vsO_Shape.Cells("Prop.Hoses66") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses51Have = pi_Hoses51Have + (vsO_Shape.Cells("Prop.Hoses51") - vsO_Shape.Cells("Prop.Hoses51") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses77Have = pi_Hoses77Have + (vsO_Shape.Cells("Prop.Hoses77") - vsO_Shape.Cells("Prop.Hoses77") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses89Have = pi_Hoses89Have + (vsO_Shape.Cells("Prop.Hoses89") - vsO_Shape.Cells("Prop.Hoses89") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses110Have = pi_Hoses110Have + (vsO_Shape.Cells("Prop.Hoses110") - vsO_Shape.Cells("Prop.Hoses110") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses150Have = pi_Hoses150Have + (vsO_Shape.Cells("Prop.Hoses150") - vsO_Shape.Cells("Prop.Hoses150") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses200Have = pi_Hoses200Have + (vsO_Shape.Cells("Prop.Hoses200") - vsO_Shape.Cells("Prop.Hoses200") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses250Have = pi_Hoses250Have + (vsO_Shape.Cells("Prop.Hoses250") - vsO_Shape.Cells("Prop.Hoses250") / vsO_Shape.Cells("Prop.Quantity"))
'                        pi_Hoses300Have = pi_Hoses300Have + (vsO_Shape.Cells("Prop.Hoses300") - vsO_Shape.Cells("Prop.Hoses300") / vsO_Shape.Cells("Prop.Quantity"))
                        
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
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipANR 'АНР
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAPP 'АПП
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAL 'АЛ
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;ALCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKP 'АКП
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;AKPCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASO 'АСО
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAT 'АТ
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipADU 'АДУ
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPNS 'ПНС
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAA 'AA
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAV 'АВ
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKT 'АКТ
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAP 'АП
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAGVT 'АГВТ
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAG 'АГТ
'                                pi_TargetedPAHave = pi_TargetedPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAGDZS 'АГДЗС
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_AGCount = pi_AGCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "TargetedPAHave;AGCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPK 'ПКС
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipPozhLab 'ЛБ
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASA 'АСА
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipASH 'АШ
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAR 'АР
'                                pi_SpecialPAHave = pi_SpecialPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SpecialPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                                
                            Case Is = indexPers.ipACL 'АЦЛ
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_ALCount = pi_ALCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;ALCount;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAKP 'АЦКП
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_AKPCount = pi_AKPCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_ACCount = pi_ACCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;AKPCount;ACCount;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAPP 'АПП
'                                pi_MainPAHave = pi_MainPAHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MainPAHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                        
                        '---Прочая пожарная техника----------
                            Case Is = indexPers.ipPoezd 'Поезда
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipSamolet 'Самолет
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipHydroSamolet 'Самолет-амфибия
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipVertolet 'Вертолет
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMotoPump 'Мотопомпы
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipKorabl 'Корабль
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipKater 'Катер
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMashNaGusenicah 'Машины на гусеничном ходу
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipTank 'Танки
'                                pi_OtherTechincsHave = pi_OtherTechincsHave + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "OtherTechincsHave;TotalTechCount;TotalFireCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMVD 'МВД
'                                pi_MVDCount = pi_MVDCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "MVDCount;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipMashiniMinzdrav 'Минздрав
'                                pi_MZdravCount = pi_MZdravCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                                .Raise "SMPCount;TotalTechCount", CellVal(shp, "Prop.Quantity") - 1
                            Case Is = indexPers.ipAvtoObschee 'Общее обозначение
                                If CellVal(shp, "Prop.RSCHS", visUnitsString) = "Входит" _
                                  Then .Raise "TotalTechCount", CellVal(shp, "Prop.Quantity") - 1 ' pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
                            Case Is = indexPers.ipPAObscheeOboznachenie 'Общее обозначение
'                                pi_TechTotalCount = pi_TechTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
'                                pi_FireTotalCount = pi_FireTotalCount + (vsO_Shape.Cells("Prop.Quantity") - 1)
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
   'Суммируем имеющиеся рукава
    With elements
        .SetVal "HosesHave", .Sum("Hoses38Have;Hoses51Have;Hoses66Have;Hoses77Have;Hoses89Have;Hoses110Have;Hoses150Have;Hoses200Have;Hoses250Have;Hoses300Have")
        .Raise "PersonnelNeed", .Result("GDZSMansRezCountNeed")
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
    End With
    
    

'---Устанавливаем разницу значений начала процедуры и конца
'tempTime2 = Time()
'MsgBox tempTime2 - tempTime1
'MsgBox "123"
'Debug.Print tempTime2 - tempTime1
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
'            If con.FromSheet.Cells("User.IndexPers").Result(visNumber) = 150 Then haveLabel = True
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
'            If con.FromSheet.Cells("User.IndexPers").Result(visNumber) = 151 Then havePos = True
            havePos = True
        End If
    Next con

    If elements.Result("AllHosesWithPos") Then elements.SetVal "AllHosesWithPos", havePos
    IsHoseShapeHavePos = havePos
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
