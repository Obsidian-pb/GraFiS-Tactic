Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long)
'Управляющая процедура импорта ТТХ аппаратов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer

On Error GoTo EX

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения значений ТТХ для текущей фигуры
    Select Case IndexPers
        Case Is = 46
            GetValuesOfCellsFromTable ShpIndex, "ДАСВ"
        Case Is = 90
            GetValuesOfCellsFromTable ShpIndex, "ДАСК"
        Case Is = 49 'Дымососы
            FogRMKGetValuesOfCellsFromTable ShpIndex, "Дымососы"
    
    End Select

Exit Sub
EX:
    SaveLog Err, "GetTTH", CStr(ShpIndex)
End Sub



