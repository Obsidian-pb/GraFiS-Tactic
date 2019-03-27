Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long)
'Управляющая процедура импорта ТТХ автомобилей
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
Select Case IndexPers
    Case Is = 1
        GetValuesOfCellsFromTable ShpIndex, "З_Автоцистерны"
    Case Is = 2
        GetValuesOfCellsFromTable ShpIndex, "З_АНР"
    Case Is = 3
        GetValuesOfCellsFromTable ShpIndex, "З_АЛ"
    Case Is = 4
        GetValuesOfCellsFromTable ShpIndex, "З_АКП"
    Case Is = 5
        GetValuesOfCellsFromTable ShpIndex, "З_АСО"
    Case Is = 6
        GetValuesOfCellsFromTable ShpIndex, "З_АТ"
    Case Is = 7
        GetValuesOfCellsFromTable ShpIndex, "З_АД"
    Case Is = 8
        GetValuesOfCellsFromTable ShpIndex, "З_ПНС"
    Case Is = 9
        GetValuesOfCellsFromTable ShpIndex, "З_АА"
    Case Is = 10
        GetValuesOfCellsFromTable ShpIndex, "З_АВ"
    Case Is = 11
        GetValuesOfCellsFromTable ShpIndex, "З_АКТ"
    Case Is = 12
        GetValuesOfCellsFromTable ShpIndex, "З_АП"
    Case Is = 13
        GetValuesOfCellsFromTable ShpIndex, "З_АГВТ"
    Case Is = 14
        GetValuesOfCellsFromTable ShpIndex, "З_АГТ"
    Case Is = 15
        GetValuesOfCellsFromTable ShpIndex, "З_АГДЗС"
    Case Is = 16
        GetValuesOfCellsFromTable ShpIndex, "З_ПКС"
    Case Is = 17
        GetValuesOfCellsFromTable ShpIndex, "З_ЛБ"
    Case Is = 18
        GetValuesOfCellsFromTable ShpIndex, "З_АСА"
    Case Is = 19
        GetValuesOfCellsFromTable ShpIndex, "З_АШ"
    Case Is = 20
        GetValuesOfCellsFromTable ShpIndex, "З_АР"
    Case Is = 161
        GetValuesOfCellsFromTable ShpIndex, "З_АЦЛ"
    Case Is = 162
        GetValuesOfCellsFromTable ShpIndex, "З_АЦКП"
    Case Is = 163
        GetValuesOfCellsFromTable ShpIndex, "З_АПП"
        
        
        
End Select



End Sub


