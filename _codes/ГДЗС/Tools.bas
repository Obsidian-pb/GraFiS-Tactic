Attribute VB_Name = "Tools"

Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'Процедура импорта данных о ТТХ любой фигуры c "Набором" из базы данных Signs
Dim dbs As DAO.Database, rsAD As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, AirDeviceModel As String
Dim i, k As Integer 'Индексы итерации

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
    On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    AirDeviceModel = ShpObj.Cells("Prop.AirDevice").ResultStr(visUnitsString)
'    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.AirDevice.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "Модели в наборе " & PASet & " отсутствуют!", vbInformation
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & AirDeviceModel & "'"
    
'---Создаем соединение с БД Signs
    pth = ThisDocument.path & "Signs.fdb"
'    Set dbs = DBEngine.OpenDatabase(pth)
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) 'Создание набора записей

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rsAD
        .FindFirst Criteria
'MsgBox .RecordCount
    '---Перебираем все строки фигуры
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---Перебираем все поля набора записей
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---Присваиеваем ячейкам активной фигуры значения в соответсвии с их ворматом в БД
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  'Текст
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   'Число
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = False
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = True
                    End If
                    
                End If
                
            Next k
        Next i

    End With

'---Закрываем соединение с БД
    Set rsAD = Nothing
    Set dbs = Nothing

Exit Sub

'---В случае ошибки открытия слишком большого количества таблиц, заканчиваем процедуру
Tail:
    MsgBox Err.description
    Set rsAD = Nothing
    Set dbs = Nothing
    SaveLog Err, "GetValuesOfCellsFromTable", "Tablename: " & TableName
End Sub


Public Sub FogRMKGetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'Процедура импорта данных о ТТХ дымососов
Dim dbs As DAO.Database, rsAD As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, FogRMKModel As String
Dim i, k As Integer 'Индексы итерации

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
    On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    FogRMKModel = ShpObj.Cells("Prop.FogRMK").ResultStr(visUnitsString)
'    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.FogRMK.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "Модели в наборе " & PASet & " отсутствуют!", vbInformation
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & FogRMKModel & "'"
    
'---Создаем соединение с БД Signs
    pth = ThisDocument.path & "Signs.fdb"
'    Set dbs = DBEngine.OpenDatabase(pth)
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) 'Создание набора записей

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rsAD
        .FindFirst Criteria
'MsgBox .RecordCount
    '---Перебираем все строки фигуры
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---Перебираем все поля набора записей
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---Присваиеваем ячейкам активной фигуры значения в соответсвии с их ворматом в БД
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  'Текст
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   'Число
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = False
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
                        'ShpObj.CellsSRC(visSectionProp, i, visCustPropsInvis).FormulaU = True
                    End If
                    
                End If
                
            Next k
        Next i

    End With

'---Закрываем соединение с БД
Set rsAD = Nothing
Set dbs = Nothing

Exit Sub

'---В случае ошибки открытия слишком большого количества таблиц, заканчиваем процедуру
Tail:
    MsgBox Err.description
    Set rsAD = Nothing
    Set dbs = Nothing
    SaveLog Err, "FogRMKGetValuesOfCellsFromTable", "Tablename: " & TableName
End Sub

Public Function ListImport(TableName As String, FieldName As String) As String
'Функция получения независимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field

    On Error GoTo EX

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
    '    Set dbs = DBEngine.OpenDatabase(pth)
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
    With rst
        .MoveFirst
        Do Until .EOF
            List = List & Replace(RSField, Chr(34), "") & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport = List

Exit Function
EX:
    SaveLog Err, "ListImport", "Tablename: " & TableName
End Function

Public Function ListImport2(TableName As String, FieldName As String, FieldName2 As String, Criteria As String) As String
'Функция получения зависимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

On Error GoTo EX

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "], [" & FieldName2 & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "], [" & FieldName2 & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ') " & _
            "AND (([" & FieldName2 & "])= '" & Criteria & "');"
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
    '    Set dbs = DBEngine.OpenDatabase(pth)
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---Проверяем количество записей в наборе и если их 0 возвращаем 0
        If rst.RecordCount > 0 Then
        '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
            With rst
                .MoveFirst
                Do Until .EOF
                    List = List & Replace(RSField, Chr(34), "") & ";"
                    .MoveNext
                Loop
            End With
        Else
            'MsgBox "Модели в наборе " & PASet & " отсутствуют!", vbInformation
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport2 = List

Exit Function
EX:
    SaveLog Err, "ListImport2", "Tablename: " & TableName
End Function

Public Function GetDBEngine() As Object
'Function returns DBEngine for current Office Engine Type (DAO.DBEngine.60 or DAO.DBEngine.120)
Dim engine As Object
    On Error GoTo EX
    Set GetDBEngine = DBEngine
Exit Function
EX:
    Set GetDBEngine = CreateObject("DAO.DBEngine.120")
End Function

'-----------------------------------------Процедуры работы с фигурами----------------------------------------------
Public Sub SetCheckForAll(ShpObj As Visio.Shape, aS_CellName As String, aB_Value As Boolean)
'Процедура устанавливает новое значение для всех выбранных фигур одного типа
Dim shp As Visio.Shape
    
    'Перебираем все фигуры в выделении и если очередная фигура имеет такую же ячейку - присваиваем ей новое значение
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).Formula = aB_Value
        End If
    Next shp
    
End Sub

Public Sub SetInnerCaptionForAll(ShpObj As Visio.Shape, aS_CellName As String)
'Процедура устанавливает содержимое для произвольной ячейки
Dim v_Str As String
Dim shp As Visio.Shape

    v_Str = InputBox("Укажите новую подпись", "Изменение содержимого")
    'Перебираем все фигуры в выделении и если очередная фигура имеет такую же ячейку - присваиваем ей новое значение
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).FormulaU = """" & v_Str & """"
        End If
    Next shp
End Sub

Public Sub MoveMeFront(ShpObj As Visio.Shape)
'Прока перемещает фигуру вперед
    ShpObj.BringToFront
End Sub

Public Function IsFirstDrop(ShpObj As Visio.Shape)
'Функция проверяет вброшенали фигура впервые и если вброшена впервые добавляет строчку свойства User.InPage
    If ShpObj.CellExists("User.InPage", 0) = 0 Then
        Dim newRowIndex As Integer
        newRowIndex = ShpObj.AddNamedRow(visSectionUser, "InPage", visRowUser)
        ShpObj.CellsSRC(visSectionUser, newRowIndex, 0).Formula = 1
        ShpObj.CellsSRC(visSectionUser, newRowIndex, visUserPrompt).FormulaU = """+"""
        
        IsFirstDrop = True
    Else
        IsFirstDrop = False
    End If
End Function

'-----------------------------------------Функции проверки привязки данных----------------------------------------------
Public Function IsShapeLinkedToDataAndDropFirst(ByRef shp As Visio.Shape) As Boolean
IsShapeLinkedToDataAndDropFirst = False

    If IsShapeLinked(shp) And shp.CellExists("User.InPage", 0) = False Then
        IsShapeLinkedToDataAndDropFirst = True
        Exit Function
    End If
End Function
Public Function IsShapeLinked(ByRef shp As Visio.Shape) As Boolean
Dim rst As Visio.DataRecordset
Dim propIndex As Integer
    
    IsShapeLinked = False
    
    For Each rst In Application.ActiveDocument.DataRecordsets
        For propIndex = 0 To shp.RowCount(visSectionProp)
            If shp.IsCustomPropertyLinked(rst.ID, propIndex) Then
                IsShapeLinked = True
                Exit Function
            End If
        Next propIndex
    Next rst
End Function


'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
'Прока сохранения лога программы
Dim errString As String
Const d = " | "

'---Открываем файл лога (если его нет - создаем)
    Open ThisDocument.path & "/Log.txt" For Append As #1
    
'---Формируем строку записи об ошибке (Дата | ОС | Path | APPDATA
    errString = Now & d & Environ("OS") & d & Environ("HOMEPATH") & d & Environ("APPDATA") & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub


