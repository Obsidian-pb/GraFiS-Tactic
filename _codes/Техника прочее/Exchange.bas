Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(shp As Visio.Shape)
'Управляющая процедура импорта ТТХ автомобилей
'---Объявляем переменные
'Dim shp As Visio.Shape
Dim IndexPers As Integer

    On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
'    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
    Select Case IndexPers
        Case Is = 73 'Машины на гусеничном ходу
            GetValuesOfCellsFromTable shp, "[З_Гусеничные машины]"
        Case Is = 74 'Танки
            GetValuesOfCellsFromTable shp, "[З_Гусеничные машины]"
        Case Is = 30 'Корабль
            GetValuesOfCellsFromTableSea shp, "З_Суда"
        Case Is = 31 'Катер
            GetValuesOfCellsFromTableSea shp, "З_Суда"
        Case Is = 24 'Поезда
            GetValuesOfCellsFromTableTrain shp, "З_Поезда"
        Case Is = 28 'Мотопомпы
            GetValuesOfCellsFromTable shp, "З_Мотопомпы"
        Case Is = 25 'Самолет
            GetValuesOfCellsFromTable shp, "З_Самолеты"
        Case Is = 26 'Самолет-амфибия
            GetValuesOfCellsFromTable shp, "З_Самолеты"
        Case Is = 27 'Вертолет
            GetValuesOfCellsFromTable shp, "З_Вертолеты"
            
            
    End Select

Exit Sub
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetTTH"
End Sub



