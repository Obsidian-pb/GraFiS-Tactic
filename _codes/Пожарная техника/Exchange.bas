Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(shp As Visio.Shape)
'Управляющая процедура импорта ТТХ автомобилей
'---Объявляем переменные
Dim IndexPers As Integer

'---Проверяем к какой именно фигуре относится данная ячейка
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
Select Case IndexPers
    Case Is = 1
        GetValuesOfCellsFromTable shp, "З_Автоцистерны"
    Case Is = 2
        GetValuesOfCellsFromTable shp, "З_АНР"
    Case Is = 3
        GetValuesOfCellsFromTable shp, "З_АЛ"
    Case Is = 4
        GetValuesOfCellsFromTable shp, "З_АКП"
    Case Is = 5
        GetValuesOfCellsFromTable shp, "З_АСО"
    Case Is = 6
        GetValuesOfCellsFromTable shp, "З_АТ"
    Case Is = 7
        GetValuesOfCellsFromTable shp, "З_АД"
    Case Is = 8
        GetValuesOfCellsFromTable shp, "З_ПНС"
    Case Is = 9
        GetValuesOfCellsFromTable shp, "З_АА"
    Case Is = 10
        GetValuesOfCellsFromTable shp, "З_АВ"
    Case Is = 11
        GetValuesOfCellsFromTable shp, "З_АКТ"
    Case Is = 12
        GetValuesOfCellsFromTable shp, "З_АП"
    Case Is = 13
        GetValuesOfCellsFromTable shp, "З_АГВТ"
    Case Is = 14
        GetValuesOfCellsFromTable shp, "З_АГТ"
    Case Is = 15
        GetValuesOfCellsFromTable shp, "З_АГДЗС"
    Case Is = 16
        GetValuesOfCellsFromTable shp, "З_ПКС"
    Case Is = 17
        GetValuesOfCellsFromTable shp, "З_ЛБ"
    Case Is = 18
        GetValuesOfCellsFromTable shp, "З_АСА"
    Case Is = 19
        GetValuesOfCellsFromTable shp, "З_АШ"
    Case Is = 20
        GetValuesOfCellsFromTable shp, "З_АР"
    Case Is = 161
        GetValuesOfCellsFromTable shp, "З_АЦЛ"
    Case Is = 162
        GetValuesOfCellsFromTable shp, "З_АЦКП"
    Case Is = 163
        GetValuesOfCellsFromTable shp, "З_АПП"
        
        
        
End Select



End Sub


