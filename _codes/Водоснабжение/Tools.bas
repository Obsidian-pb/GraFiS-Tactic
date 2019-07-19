Attribute VB_Name = "Tools"
Public Function ListImport(TableName As String, FieldName As String) As String
'Функция получения независимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field

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
            List = List & RSField & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport = List

Set dbs = Nothing
Set rst = Nothing
End Function

Public Function ListImport2(TableName As String, FieldName As String, Criteria As String) As String
'Функция получения зависимого списка из базы данных
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=' ') " & _
            "And " & Criteria & _
        " GROUP BY [" & FieldName & "]; "
        
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
                    List = List & RSField & ";"
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
End Function

Public Function ListImportNum(TableName As String, FieldName As String, Criteria As String) As String
'Функция получения зависимого списка из базы данных (Для цифровых значений)
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As DAO.Field, RSField2 As DAO.Field

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]=0) " & _
            "And " & Criteria & _
        " GROUP BY [" & FieldName & "]; "
        
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
                    List = List & RSField & ";"
                    .MoveNext
                Loop
            End With
        Else
            'MsgBox "Модели в наборе " & PASet & " отсутствуют!", vbInformation
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImportNum = List

Set dbs = Nothing
Set rst = Nothing
End Function

Public Function ValueImportStr(TableName As String, FieldName As String, Criteria As String) As String
'Процедура получения значения произвольного поля таблицы соответствующего другим полям этой же таблицы
Dim dbs As Database, rst As Recordset
Dim pth As String
Dim SQLQuery As String
Dim RSField As DAO.Field
Dim ValOfSerch As String

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
        With rst
            .MoveFirst
            ValOfSerch = RSField
        End With
        ValOfSerch = Chr(34) & ValOfSerch & Chr(34)

ValueImportStr = ValOfSerch

Set dbs = Nothing
Set rst = Nothing
End Function


Public Function ValueImportSng(TableName As String, FieldName As String, Criteria As String) As Single
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
        "WHERE ([" & FieldName & "] Is Not Null Or Not [" & FieldName & "]= 0) " & _
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
    ValueImportSng = "0"

    Set dbs = Nothing
    Set rst = Nothing
End Function



Public Function CD_MasterExists(MasterName As String) As Boolean
'Функция проверки наличия мастера в активном документе
Dim i As Integer

    For i = 1 To Application.ActiveDocument.Masters.Count
        If Application.ActiveDocument.Masters(i).Name = MasterName Then
            CD_MasterExists = True
            Exit Function
        End If
    Next i
    
    CD_MasterExists = False

End Function

Public Sub MasterImportSub(DocName As String, MasterName As String)
'Процедура импорта мастера в соответствии с именем
Dim mstr As Visio.Master

    If Not CD_MasterExists(MasterName) Then
        Set mstr = Application.Documents(DocName).Masters(MasterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

End Sub

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

'--------------------------------Проверка возможности обращения фигур---------------------
Public Function IsSelectedOneShape(ShowMessage As Boolean) As Boolean
'---Проверяем выбран ли какой либо одиносный объект
    If Not Application.ActiveWindow.Selection.Count = 1 Then
        If ShowMessage Then MsgBox "Не выбрана ни одна фигура или выбрано больше одной фигуры", vbInformation
        IsSelectedOneShape = False
        Exit Function
    End If
IsSelectedOneShape = True
End Function
Public Function IsHavingUserSection(ShowMessage As Boolean) As Boolean
'---Проверяем, не является ли выбранная фигура уже фигурой с назначенными свойствами
    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
        If ShowMessage Then MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в зону горения", vbInformation
        IsHavingUserSection = False
        Exit Function
    End If
IsHavingUserSection = True
End Function
Public Function IsSquare(ShowMessage As Boolean) As Boolean
'---Проверяем Является ли выбранная фигура площадью
    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
        If ShowMessage Then MsgBox "Выбранная фигура не имеет площади!", vbInformation
        IsSquare = False
        Exit Function
    End If
IsSquare = True
End Function

Public Sub ShowCommonData(shp As Visio.Shape)
'Показываем общие сведения о водоисточнике
Dim common As String

    common = shp.Cells("Prop.Common").ResultStr(0)
    
    f_INPPV_CommonData.ShowData common
'    f_INPPV_CommonData.Show

    
'    MsgBox common
End Sub

'Public Function CheckSquareShape(ShowMessage As Boolean) As Boolean
''Функция возвращает Истина, если фигуру можно преобразовать в фигуру c площадью, Ложь - если нет
'
'    On Error GoTo EX
''---Проверяем выбран ли какой либо объект
'    If Application.ActiveWindow.Selection.Count < 1 Then
''        MsgBox "Не выбрана ни одна фигура!", vbInformation
'        CheckSquareShape = False
'        Exit Function
'    End If
'
''---Проверяем, не является ли выбранная фигура уже фигурой с назначенными свойствами
'    If Application.ActiveWindow.Selection(1).RowCount(visSectionUser) > 0 Then
'        MsgBox "Выбранная фигура уже имеет специальные свойства и не может быть обращена в зону горения", vbInformation
'        CheckSquareShape = False
'        Exit Function
'    End If
'
''---Проверяем Является ли выбранная фигура площадью
'    If Application.ActiveWindow.Selection(1).AreaIU = 0 Then
'        MsgBox "Выбранная фигура не имеет площади!", vbInformation
'        CheckSquareShape = False
'        Exit Function
'    End If
'
'CheckSquareShape = True
'Exit Function
'EX:
'    SaveLog Err, "CheckSquareShape"
'    CheckSquareShape = False
'End Function

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

'-----------TESTS-----------------
'Public Function ListImport222() As String
''Функция получения независимого списка из базы данных
'
'Dim dbs As Database, rst As Recordset
'Dim pth As String
'Dim ws As Workspace
''Dim SQLQuery As String
''Dim List As String
''Dim RSField As DAO.Field
'
''---Определяем набор записей
'    '---Определяем запрос SQL для отбора записей из базы данных
''        SQLQuery = "SELECT [" & FieldName & "] " & _
''        "FROM [" & TableName & "] " & _
''        "GROUP BY [" & FieldName & "] " & _
''        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
'
'    '---Создаем набор записей для получения списка
'        pth = "D:\Signs.fdb"
'        Set ws = DBEngine.CreateWorkspace("newWS", "admin", "", dbUseJet)
''        Set ws = CreateWorkspace("", "admin", "", dbUseJet)
'        Dim sss As DAO.DBEngine
'        Set sss = CreateObject("DAO.DBEngine")
''        Set dbs = sss.OpenDatabase(pth)
''        Set rst = dbs.CreateQueryDef("", SQLQuery).OpenRecordset(dbOpenDynaset)  'Создание набора записей
''        Set RSField = rst.Fields(FieldName)
'
'    '---Ищем необходимую запись в наборе данных и по ней создаем набор значений для списка для заданных параметров
''    With rst
''        .MoveFirst
''        Do Until .EOF
''            List = List & RSField & ";"
''            .MoveNext
''        Loop
''    End With
''    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
''ListImport = List
'
''Set dbs = Nothing
'Set rst = Nothing
'End Function
