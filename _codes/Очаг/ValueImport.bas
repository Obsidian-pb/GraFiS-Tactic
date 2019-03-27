Attribute VB_Name = "ValueImport"
'------------------------Модуль для процедур импорта значений ячеек-------------------
'------------------------Блок Значений ячеек------------------------------------------
Public Sub GetFactorsByDescription(ShpIndex As Long)
'Процедура импорта данных о интенсивностях и линейной скорости из базы данных Signs по описанию пожара
Dim dbsE As Database, rsType As Recordset
Dim pth As String
Dim Critria As String, Categorie As String, description As String, IntenseW As Single, speed As Single
Dim shp As Visio.Shape

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
    On Error GoTo Tail

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    Categorie = shp.Cells("Prop.FireCategorie").ResultStr(visUnitsString)
    description = shp.Cells("Prop.FireDescription").ResultStr(visUnitsString)
    Criteria = "[Категория] = '" & Categorie & "' And [Описание] = '" & description & "'"
    
'---Создаем соединение с БД Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbsE = GetDBEngine.OpenDatabase(pth)
    Set rsType = dbsE.OpenRecordset("З_Интенсивности", dbOpenDynaset) 'Создание набора записей

'---Ищем необходимую запись в наборе данных и по ней определяем интенсивность для заданных параметров
    With rsType
        .FindFirst Criteria
        If ![ИнтенсивностьПоВодеРасч] > 0 Then 'Если значения интенсивности подачи воды в БД нет
            Intense = ![ИнтенсивностьПоВодеРасч]
        Else
            MsgBox "Расчетное значение интенсивности подачи воды для данного описания в базе данных отсутствует! " & _
                "Поэтому по умолчанию будет присовено значение 0л/с*м.кв.."
            Intense = 0
        End If
        
        If ![СкоростьРасч] > 0 Then 'Если значения скорости в БД нет
            speed = ![СкоростьРасч]
        Else
            MsgBox "Расчетное значение линейной скорости распространения огня для данного описания в базе данных отсутствует! " & _
                "Поэтому по умолчанию будет присовено значение 0м/мин."
            speed = 0
        End If
    End With
    
'---Присваиваем полученные значения ячейкам
        shp.Cells("Prop.WaterIntense").FormulaU = str(Intense)
        shp.Cells("Prop.FireSpeedLine").FormulaU = str(speed)
    
'---Закрываем соединение с БД
rsType.Close
dbsE.Close
Set dbs = Nothing

Exit Sub

'---В случае ошибки открытия слишком большого количества таблиц, заканчиваем процедуру
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "Sm_ShapeFormShow"
End Sub

