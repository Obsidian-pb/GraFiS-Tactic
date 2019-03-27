Attribute VB_Name = "Tools"

Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'Процедура импорта данных о ТТХ любой фигуры c "Набором" из базы данных Signs
Dim dbs As DAO.Database, rsPA As DAO.Recordset
Dim pth As String
Dim ShpObj As Visio.Shape
Dim Critria As String, PAModel As String, PASet As String
Dim i, k As Integer 'Индексы итерации

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            'MsgBox "Модели в наборе " & PASet & " отсутствуют!", vbInformation
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & PAModel & "' And [Набор] = '" & PASet & "'"
    
'---Создаем соединение с БД Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = GetDBEngine.OpenDatabase(pth)
    Set rsPA = dbs.OpenRecordset(TableName, dbOpenDynaset) 'Создание набора записей

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rsPA
        .FindFirst Criteria
'MsgBox .RecordCount
    '---Перебираем все строки фигуры
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---Перебираем все поля набора записей
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).value >= 0 Then
                        '---Присваиеваем ячейкам активной фигуры значения в соответсвии с их ворматом в БД
                        'MsgBox .Fields(k).Type & "? " & .Fields(k).Name
                        If .Fields(k).Type = 10 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).value & """"  'Текст
                        If .Fields(k).Type = 6 Or .Fields(k).Type = 4 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).value)   'Число
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
Set rsPA = Nothing
Set dbs = Nothing

Exit Sub

'---В случае ошибки открытия слишком большого количества таблиц, заканчиваем процедуру
Tail:
'MsgBox Err.Description
    Set rsPA = Nothing
    Set dbs = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "GetValuesOfCellsFromTable"

End Sub

'----------------------------------------------------------------------------------------------

Public Function ListImport(TableName As String, FieldName As String) As String
'Функция получения независимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field

    On Error GoTo Tail

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
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

Set dbs = Nothing
Set rst = Nothing
Exit Function
Tail:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ListImport"
End Function


Public Function ListImportInt(TableName As String, FieldName As String) As String
'Функция получения независимого списка из базы данных (для числовых значений)
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
        "HAVING (([" & FieldName & "]) Is Not Null);"
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
    With rst
        .MoveFirst
        Do Until .EOF
            List = List & RSField & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImportInt = List

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ListImportInt"
End Function


Public Function ListImport2(TableName As String, FieldName As String, Criteria As String) As String
'Функция получения зависимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

    On Error GoTo EX

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            "And " & Criteria & _
        "GROUP BY [" & FieldName & "]; "
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
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

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ListImport2"
End Function


Public Function ValueImportStr(TableName As String, FieldName As String, Criteria As String) As String
'Процедура получения значения произвольного поля таблицы соответствующего другим полям этой же таблицы
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim RSField As DAO.Field
Dim ValOfSerch As String

    On Error GoTo EX

'---Определяем запись с соответствующи параметром
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            " And " & Criteria & "; "
            
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---В соответствии с полученной записью возвращаем значение искомого поля
        If rst.RecordCount > 0 Then
            With rst
                .MoveFirst
                ValOfSerch = RSField
            End With
        Else
            ValOfSerch = ""
        End If
        
        ValOfSerch = Chr(34) & ValOfSerch & Chr(34)

ValueImportStr = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ValueImportStr"
End Function


Public Function ValueImportSng(TableName As String, FieldName As String, Criteria As String) As Single
'Процедура получения значения произвольного поля таблицы соответствующего другим полям этой же таблицы
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim RSField As DAO.Field
Dim ValOfSerch As String

    On Error GoTo EX

'---Определяем запись с соответствующи параметром
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            " And " & Criteria & "; "
            
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---В соответствии с полученной записью возвращаем значение искомого поля
        If rst.RecordCount > 0 Then
            With rst
                .MoveFirst
                ValOfSerch = RSField
            End With
        Else
            ValueImportSng = 0
            Set dbs = Nothing
            Set rst = Nothing
            Exit Function
        End If

ValueImportSng = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ValueImportSng"
End Function

Public Function ValueImportSngStr(TableName As String, FieldName As String, Criteria As String) As Single 'При числовом критерии
'Процедура получения значения произвольного поля таблицы соответствующего другим полям этой же таблицы
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim RSField As DAO.Field
Dim ValOfSerch As Single

    On Error GoTo EX

'---Определяем запись с соответствующи параметром
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
            " And " & Criteria & "; "
            
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = GetDBEngine.OpenDatabase(pth)
        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
        Set RSField = rst.Fields(FieldName)
        
    '---В соответствии с полученной записью возвращаем значение искомого поля
        If rst.RecordCount > 0 Then
            With rst
                .MoveFirst
                ValOfSerch = RSField
            End With
        Else
            ValueImportSngStr = 0
            Set dbs = Nothing
            Set rst = Nothing
            Exit Function
        End If

ValueImportSngStr = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    Set dbs = Nothing
    Set rst = Nothing
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ValueImportSngStr"
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

