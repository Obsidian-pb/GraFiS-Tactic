Attribute VB_Name = "Tools"
Option Explicit



Public Sub GetValuesOfCellsFromTable(ShpIndex As Long, TableName As String)
'Процедура импорта данных о ТТХ любой фигуры c "Набором" из базы данных Signs
Dim dbs As Object, rst As Object
Dim pth As String
Dim ShpObj As Visio.Shape
Dim SQL As String, Criteria As String, PAModel As String, PASet As String
Dim i, k As Integer 'Индексы итерации
Dim fieldType As Integer

'---Определяем действие в случае ошибки открытия слишком большого количества таблиц
On Error GoTo Tail

'---Определяем фигуру относительно которой выполняется действие
    Set ShpObj = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Определяем критерии поиска записи в наборе данных
    PAModel = ShpObj.Cells("Prop.Model").ResultStr(visUnitsString)
    PASet = ShpObj.Cells("Prop.Set").ResultStr(visUnitsString)
        If ShpObj.Cells("Prop.Model.Format").ResultStr(visUnitsString) = "" Then
            Exit Sub 'Если в ячейке "Модель" пустое значение - процедура прекращается
        End If
    Criteria = "[Модель] = '" & PAModel & "' And [Набор] = '" & PASet & "'"
    
'---Создаем соединение с БД Signs
    pth = ThisDocument.path & "Signs.fdb"
    Set dbs = CreateObject("ADODB.Connection")
    dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
    dbs.Open
    Set rst = CreateObject("ADODB.Recordset")
    SQL = "SELECT * From " & TableName
    rst.Open SQL, dbs, 3, 1

    

'---Ищем необходимую запись в наборе данных и по ней определяем ТТХ ПА для заданных параметров
    With rst
        .Filter = Criteria
        If .RecordCount > 0 Then
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
        End If
    End With

'---Закрываем соединение с БД
Set rst = Nothing
Set dbs = Nothing

Exit Sub

'---В случае ошибки открытия слишком большого количества таблиц, заканчиваем процедуру
Tail:
    MsgBox Err.description
    Set rst = Nothing
    Set dbs = Nothing
    SaveLog Err, "GetValuesOfCellsFromTable"

End Sub

'----------------------------------------------------------------------------------------------
Public Function ListImport(TableName As String, FieldName As String) As String
'Функция получения независимого списка из базы данных
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim List As String
Dim RSField As Object

    On Error GoTo EX

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
            List = List & Replace(RSField, Chr(34), "") & ";"
            .MoveNext
        Loop
    End With
    List = Chr(34) & Left(List, Len(List) - 1) & Chr(34)
    ListImport = List

Set dbs = Nothing
Set rst = Nothing
Exit Function
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ListImport"
    ListImport = Chr(34) & " " & Chr(34)
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

    On Error GoTo EX

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
                    List = List & Replace(RSField, Chr(34), "") & ";"
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
Exit Function
EX:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "ListImport2"
    ListImport2 = Chr(34) & " " & Chr(34)
    Set dbs = Nothing
    Set rst = Nothing
End Function

Public Function ValueImportStr(TableName As String, FieldName As String, Criteria As String) As String
'Процедура получения значения произвольного поля таблицы соответствующего другим полям этой же таблицы
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As String

'---Определяем запись с соответствующи параметром
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
            " And " & Criteria & "; "
            
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
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
Dim dbs As Object, rst As Object
Dim pth As String
Dim SQLQuery As String
Dim RSField As Object
Dim ValOfSerch As Single

'---Определяем запись с соответствующи параметром
    '---Определяем запрос SQL для отбора записей из базы данных
        SQLQuery = "SELECT [" & FieldName & "]" & _
        "FROM [" & TableName & "] " & _
        "WHERE [" & FieldName & "] Is Not Null " & _
            " And " & Criteria & "; "
            
    '---Создаем набор записей для получения списка
        pth = ThisDocument.path & "Signs.fdb"
        Set dbs = CreateObject("ADODB.Connection")
        dbs = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & pth & ";Uid=Admin;Pwd=;"
        dbs.Open
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open SQLQuery, dbs, 3, 1
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
End Function






Public Function CD_MasterExists(masterName As String) As Boolean
'Функция проверки наличия мастера в активном документе
Dim i As Integer

For i = 1 To Application.ActiveDocument.Masters.Count
    If Application.ActiveDocument.Masters(i).Name = masterName Then
        CD_MasterExists = True
        Exit Function
    End If
Next i

CD_MasterExists = False

End Function

Public Sub MasterImportSub(masterName As String)
'Процедура импорта мастера в соответствии с именем
Dim mstr As Visio.Master

    If Not CD_MasterExists(masterName) Then
        Set mstr = ThisDocument.Masters(masterName)
        Application.ActiveDocument.Masters.Drop mstr, 0, 0
    End If

End Sub


Public Sub SetFormulaForAll(ShpObj As Visio.Shape, ByVal aS_CellName As String, ByVal aS_NewFormula As String)
'Процедура устанавливает содержимое для произвольной ячейки
Dim shp As Visio.Shape

    'Перебираем все фигуры в выделении и если очередная фигура имеет такую же ячейку - присваиваем ей новое значение
    For Each shp In Application.ActiveWindow.Selection
        If shp.CellExists(aS_CellName, 0) = True Then
            shp.Cells(aS_CellName).FormulaU = """" & aS_NewFormula & """"
        End If
    Next shp
End Sub

Public Sub MoveMeFront(ShpObj As Visio.Shape)
'Прока перемещает фигуру вперед
    ShpObj.BringToFront
End Sub

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
        IsHavingUserSection = True
        Exit Function
    End If
IsHavingUserSection = False
End Function
Public Function IsSquare(ShowMessage As Boolean) As Boolean
'---Проверяем Является ли выбранная фигура площадью
    If Application.ActiveWindow.Selection(1).AreaIU > 0 Then
        If ShowMessage Then MsgBox "Выбранная фигура не является фигурой линии!", vbInformation
        IsSquare = True
        Exit Function
    End If
IsSquare = False
End Function
Public Function ClickAndOnSameButton(ClickedButtonName As String) As Boolean
'---Проверяем не нажата ли та же кнопка, что активна в данный момент
'---Истина если нажата и активна одна и та же кнопка
'---Ложь, если это разные кнопки или не нажато ни одной кнопки
Dim v_Cntrl As CommandBarControl
    
'---Включаем обработку ошибок
    On Error GoTo Tail

'---Проверяем нажата ли указанная кнопка или другая
    For Each v_Cntrl In Application.CommandBars("Превращения").Controls
        If v_Cntrl.State = msoButtonDown And v_Cntrl.Caption = ClickedButtonName Then
            ClickAndOnSameButton = True
            Exit Function
        End If
    Next v_Cntrl
    
    ClickAndOnSameButton = False
Exit Function
Tail:
    ClickAndOnSameButton = False
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу."
    SaveLog Err, "NoButtonsON"
End Function

'--------------------------------Разное----------------------------------------------------------------
Public Function Index(ByVal str As String, ByVal stringArray As String, ByVal delimiter) As Integer
'Тут все и так понятно
Dim arr() As String
Dim i As Integer
    
    arr = Split(stringArray, delimiter)
    
    i = 0
    For i = 0 To UBound(arr)
        If arr(i) = str Then
            Index = i
            Exit Function
        End If
    Next i
End Function

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

'-------------------Инструменты проверки фигур
Public Function CellVal(ByRef shp As Visio.Shape, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber) As Variant
'Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
    
    On Error GoTo EX
    
    If shp.CellExists(cellName, 0) Then
        Select Case dataType
            Case Is = visNumber
                CellVal = shp.Cells(cellName).Result(dataType)
            Case Is = visUnitsString
                CellVal = shp.Cells(cellName).ResultStr(dataType)
            Case Is = visDate
                CellVal = shp.Cells(cellName).Result(dataType)
        End Select
    Else
        CellVal = 0
    End If
    
    
Exit Function
EX:
    CellVal = 0
End Function

Public Function IsGFSShape(ByRef shp As Visio.Shape, Optional ByVal useManeure As Boolean = True) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС
Dim i As Integer
    
    'Проверяем, является ли фигура фигурой ГраФиС
    If useManeure Then      'Если нужно учитывать проверку на маневр
        If shp.CellExists("User.IndexPers", 0) = True Then
            'Если имеется ячейка опции Маневра и ее значение показывает, что
            If shp.CellExists("Actions.MainManeure", 0) = True Then
                If shp.Cells("Actions.MainManeure.Checked").Result(visNumber) = 0 Then
                    IsGFSShape = True       'Фигура ГраФиС и не маневренная
                Else
                    IsGFSShape = False      'Фигура ГраФиС и маневренная
                End If
            Else
                IsGFSShape = True       'Фигура ГраФиС и не имеет ячейки Маневр
            End If
        Else
            IsGFSShape = False      'Фигура не ГраФиС
        End If
    Else                    'если не нужно учитывать проверку на маневр
        IsGFSShape = shp.CellExists("User.IndexPers", 0)
    End If

End Function

Public Function IsGFSShapeWithIP(ByRef shp As Visio.Shape, ByRef gfsIndexPerses As Variant, Optional needGFSChecj As Boolean = False) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС и среди переданных типов фигур ГраФиС (gfsIndexPreses) присутствует IndexPers данной фигуры
'По умолчанию предполагается что переданная фигура уже проверена на то, относится ли она к фигурам ГраФиС. В случае, если у фигуры нет ячейки User.IndexPers _
'обработчик ошибки указывает функции вернуть False
'Пример использования: IsGFSShapeWithIP(shp, indexPers.ipPloschadPozhara)
'                 или: IsGFSShapeWithIP(shp, Array(indexPers.ipPloschadPozhara, indexPers.ipAC))
Dim i As Integer
Dim indexPers As Integer
    
    On Error GoTo EX
    
    'Если необходима предварительная проверка на отношение фигуры к ГраФиС:
    If needGFSChecj Then
        If Not IsGFSShape(shp) Then
            IsGFSShapeWithIP = False
            Exit Function
        End If
    End If
    
    'Проверяем, является ли фигура фигурой указанного типа
    indexPers = shp.Cells("User.IndexPers").Result(visNumber)
    Select Case TypeName(gfsIndexPerses)
        Case Is = "Long"    'Если передано единственное значение Long
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Integer"    'Если передано единственное значение Integer
            If gfsIndexPerses = indexPers Then
                IsGFSShapeWithIP = True
                Exit Function
            End If
        Case Is = "Variant()"   'Если передан массив
            For i = 0 To UBound(gfsIndexPerses)
                If gfsIndexPerses(i) = indexPers Then
                    IsGFSShapeWithIP = True
                    Exit Function
                End If
            Next i
        Case Else
            IsGFSShapeWithIP = False
    End Select

IsGFSShapeWithIP = False
Exit Function
EX:
    IsGFSShapeWithIP = False
    SaveLog Err, "m_Tools.IsGFSShapeWithIP"
End Function


