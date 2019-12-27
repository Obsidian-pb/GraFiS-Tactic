Attribute VB_Name = "Tools"
Option Explicit

Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'Процедура импорта данных о ТТХ любой фигуры c "Набором" из базы данных Signs
Dim dbs As Object, rsAD As Object
Dim pth As String
Dim ShpObj As Visio.Shape
Dim SQL As String, Criteria As String, AirDeviceModel As String
Dim i, k As Integer 'Индексы итерации
Dim fieldType As Integer

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
    On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    AirDeviceModel = ShpObj.Cells("Prop.AirDevice").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.AirDevice.Format").ResultStr(visUnitsString) = "" Then
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & AirDeviceModel & "'"
    
'---Создаем соединение с БД Signs
'    pth = ThisDocument.path & "Signs.fdb"
''    Set dbs = DBEngine.OpenDatabase(pth)
'    Set dbs = GetDBEngine.OpenDatabase(pth)
'    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) 'Создание набора записей
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = CreateObject("ADODB.Connection")
    dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
    dbs.Open
    Set rsAD = CreateObject("ADODB.Recordset")
    SQL = "SELECT * From " & TableName
    rsAD.Open SQL, dbs, 3, 1

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rsAD
        .Filter = Criteria
        .MoveFirst
    '---Перебираем все строки фигуры
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---Перебираем все поля набора записей
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---Присваиеваем ячейкам активной фигуры значения в соответсвии с их форматом в БД
                        fieldType = .Fields(k).Type
                        If fieldType = 202 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  'Текст
                        If fieldType = 2 Or fieldType = 3 Or fieldType = 4 Or fieldType = 5 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   'Число
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
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
Dim dbs As Object, rsAD As Object
Dim pth As String
Dim ShpObj As Visio.Shape
Dim SQL As String, Criteria As String, FogRMKModel As String
Dim i, k As Integer 'Индексы итерации
Dim fieldType As Integer

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
    On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    FogRMKModel = ShpObj.Cells("Prop.FogRMK").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.FogRMK.Format").ResultStr(visUnitsString) = "" Then
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & FogRMKModel & "'"
    
'---Создаем соединение с БД Signs
'    pth = ThisDocument.path & "Signs.fdb"
''    Set dbs = DBEngine.OpenDatabase(pth)
'    Set dbs = GetDBEngine.OpenDatabase(pth)
'    Set rsAD = dbs.OpenRecordset(TableName, dbOpenDynaset) 'Создание набора записей
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = CreateObject("ADODB.Connection")
    dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
    dbs.Open
    Set rsAD = CreateObject("ADODB.Recordset")
    SQL = "SELECT * From " & TableName
    rsAD.Open SQL, dbs, 3, 1

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rsAD
        .Filter = Criteria
        .MoveFirst
    '---Перебираем все строки фигуры
        For i = 0 To ShpObj.RowCount(visSectionProp) - 1
        '---Перебираем все поля набора записей
            For k = 0 To .Fields.Count - 1
                If ShpObj.CellsSRC(visSectionProp, i, visCustPropsLabel).ResultStr(Visio.visNone) = _
                    .Fields(k).Name Then
                    If .Fields(k).Value >= 0 Then
                        '---Присваиеваем ячейкам активной фигуры значения в соответсвии с их ворматом в БД
                        fieldType = .Fields(k).Type
                        If fieldType = 202 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = """" & .Fields(k).Value & """"  'Текст
                        If fieldType = 2 Or fieldType = 3 Or fieldType = 4 Or fieldType = 5 Then _
                            ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = str(.Fields(k).Value)   'Число
                    Else
                        ShpObj.CellsSRC(visSectionProp, i, visCustPropsValue).FormulaU = 0
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
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "] " & _
        "FROM [" & TableName & "] " & _
        "GROUP BY [" & FieldName & "] " & _
        "HAVING (([" & FieldName & "]) Is Not Null Or Not ([" & FieldName & "])=' ');"
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
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
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object, RSField2 As Object

'---Определяем набор записей
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
            "And " & Criteria & _
        "GROUP BY [" & FieldName & "]; "
        
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
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
            List = "0"
        End If
        List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
ListImport2 = List

Set dbs = Nothing
Set rst = Nothing
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

Public Function IsShapeGraFiSType(ByRef aO_TergetShape As Visio.Shape, ByRef arr As Variant) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС и при этом имеет индекс из указанного массива
    IsShapeGraFiSType = False
    
    If aO_TergetShape.CellExists("User.IndexPers", 0) = True And aO_TergetShape.CellExists("User.Version", 0) = True Then
        IsShapeGraFiSType = IsInArray(arr, aO_TergetShape.Cells("User.IndexPers"))
    End If
End Function

Public Function IsInArray(ByRef arr As Variant, ByVal val As Integer) As Boolean
'Функция возвращает True если значение имеется в массиве, иначе False
Dim i As Integer
    For i = 0 To UBound(arr)
        If arr(i) = val Then
            IsInArray = True
            Exit Function
        End If
    Next i
IsInArray = False
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
    errString = Now & d & Environ("OS") & d & "Visio " & Application.version & d & ThisDocument.fullName & d & eroorPosition & _
        d & error.Number & d & error.description & d & error.Source & d & eroorPosition & d & addition
    
'---Записываем в конец файла лога сведения о ошибке
    Print #1, errString
    
'---Закрываем фацл лога
    Close #1

End Sub

Public Function GetPointOnLineShape(ByRef center As c_Vector, ByRef lineShape As Visio.Shape, ByVal radiuss As Double) As c_Vector
'Находим первую точку на линии
Const pi = 3.1415                           'Число Пи
Const oneOfHundredOfPerimeter = 0.06283     'угол - одна сотая окружности в радианах
Dim pointTolerance As Double                'Точность поиска хита с линией для каждой из точек окружности
Dim checkPoint As c_Vector
Dim i As Byte
    
    pointTolerance = pi * radiuss / 100
    
    Set checkPoint = New c_Vector
    
    For i = 0 To 99
'        checkPoint.x =
    Next i
    
End Function
