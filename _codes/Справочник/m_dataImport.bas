Attribute VB_Name = "m_dataImport"
Option Compare Database

Public Sub BaseDataImport()
'Стартовая прока импорта
Dim aStr_Path As String
Dim vObj_FD  'As Application.FileDialog

'---загружаем выбираемый пользователем файл
    Set vObj_FD = Application.FileDialog(1)
    With vObj_FD
        .Filters.Clear
        .AllowMultiSelect = False
        .Filters.Add "База данных ГраФиС", "*.fdb"
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        aStr_Path = .SelectedItems(1)
    End With

    MainImportProc aStr_Path
    
End Sub

Private Sub MainImportProc(ByVal outerDBPath As String)
'Главная прока импорта данных
Dim thisDB As Database
Dim outerDB As Database
    
    'Временная заглушка
'    Exit Sub

    
    Set thisDB = CurrentDb
    Set outerDB = DBEngine.OpenDatabase(outerDBPath)
    
    'Импорт списков
    ImportDataLists thisDB, outerDB, "Наборы", "Набор"
    ImportDataLists thisDB, outerDB, "Типы двигателей", "Двигатель"
    ImportDataLists thisDB, outerDB, "Типы насосов", "Насос"
    ImportDataLists thisDB, outerDB, "Типы шасси", "Шасси"
    ImportDataLists thisDB, outerDB, "Баллоны", "Обозначение"
    
    'Испорт техники
    ImportDataTechnics thisDB, outerDB, "Автоцистерны", "Модель"
    ImportDataTechnics thisDB, outerDB, "АА", "Модель"
    ImportDataTechnics thisDB, outerDB, "АВ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АГВТ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АГДЗС", "Модель"
    ImportDataTechnics thisDB, outerDB, "АГТ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АД", "Модель"
    ImportDataTechnics thisDB, outerDB, "АКП", "Модель"
    ImportDataTechnics thisDB, outerDB, "АКТ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АЛ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АЛП", "Модель"
    ImportDataTechnics thisDB, outerDB, "АНР", "Модель"
    ImportDataTechnics thisDB, outerDB, "АП", "Модель"
    ImportDataTechnics thisDB, outerDB, "АПП", "Модель"
    ImportDataTechnics thisDB, outerDB, "АР", "Модель"
    ImportDataTechnics thisDB, outerDB, "АСА", "Модель"
    ImportDataTechnics thisDB, outerDB, "АСО", "Модель"
    ImportDataTechnics thisDB, outerDB, "АТ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АТСО", "Модель"
    ImportDataTechnics thisDB, outerDB, "АЦКП", "Модель"
    ImportDataTechnics thisDB, outerDB, "АЦЛ", "Модель"
    ImportDataTechnics thisDB, outerDB, "АШ", "Модель"
    ImportDataTechnics thisDB, outerDB, "Вертолеты", "Модель"
    ImportDataTechnics thisDB, outerDB, "Гусеничные машины", "Модель"
    ImportDataTechnics thisDB, outerDB, "Мотопомпы", "Модель"
    ImportDataTechnics thisDB, outerDB, "ПКС", "Модель"
    ImportDataTechnics thisDB, outerDB, "ПНС", "Модель"
    ImportDataTechnics thisDB, outerDB, "Поезда", "Категория"
    ImportDataTechnics thisDB, outerDB, "Самолеты", "Модель"
    ImportDataTechnics thisDB, outerDB, "Суда", "Проект"
    
    'ГДЗС
    ImportDataCommon thisDB, outerDB, "ДАСВ", "Модель"
    ImportDataCommon thisDB, outerDB, "ДАСК", "Модель"
    
    'Дымососы
    ImportDataCommon thisDB, outerDB, "Дымососы", "Модель"
    
    'Рукава
    ImportDataHoses thisDB, outerDB, "Рукава", "Материал", "Диаметр"
    
    'Водопровод
'    ImportDataWater thisDB, outerDB, "Диаметры водоводов", "Диаметр водовода", "КодВидаСети"
'    ImportDataWater thisDB, outerDB, "Водоотдача", "Напор в сети", "КодДиаметра"
    
    'Стволы
'    ImportDataCommon thisDB, outerDB, "МоделиСтволов", "Модель ствола"
'    ImportDataNozzle thisDB, outerDB, "Рукава", "Материал", "Диаметр"
    
    
    
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
'    ImportDataTechnics thisDB, outerDB, "", "Модель"
    

    MsgBox "Данные успешно импортированы!"



End Sub

Private Sub ImportDataTechnics(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "' And " & _
            "[Набор] =" & GetKeyFieldValue(dbThis, dbOuter, "Наборы", "КодНабора", "Набор", rstOuter.Fields("Набор").Value)
        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "Набор" Then
                    fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Наборы", "КодНабора", "Набор", fldOuter.Value)
                ElseIf fldOuter.Name = "Насос" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Типы насосов", "КодНасоса", "Насос", fldOuter.Value)
                    End If
                ElseIf fldOuter.Name = "Шасси" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Типы шасси", "КодШасси", "Шасси", fldOuter.Value)
                    End If
                ElseIf fldOuter.Name = "Двигатель" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Типы двигателей", "КодДвигателя", "Двигатель", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataHoses(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String, ByVal keyFieldName2 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "=" & rstOuter.Fields(keyFieldName1) & " And " & _
            "[" & keyFieldName2 & "] =" & rstOuter.Fields(keyFieldName2)
        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "Диаметр" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Диаметры рукавов", "КодДиаметра", "Диаметр рукавов", fldOuter.Value)
                    End If
                ElseIf fldOuter.Name = "Материал" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Материалы рукавов", "КодМатериалаРукавов", "Материал рукава", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataWater(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String, ByVal keyFieldName2 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "=" & rstOuter.Fields(keyFieldName1) & " And " & _
            "[" & keyFieldName2 & "] =" & rstOuter.Fields(keyFieldName2)
        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "Диаметр" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Диаметры рукавов", "КодДиаметра", "Диаметр рукавов", fldOuter.Value)
                    End If
                ElseIf fldOuter.Name = "Материал" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Материалы рукавов", "КодМатериалаРукавов", "Материал рукава", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataNozzleVariants(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String, ByVal keyFieldName2 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "' And " & _
            "[КодМоделиСтвола] =" & GetKeyFieldValue(dbThis, dbOuter, "МоделиСтволов", "КодМоделиСтвола", "Модель ствола", rstOuter.Fields("КодМоделиСтвола").Value)

        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "КодМоделиСтвола" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "МоделиСтволов", "КодМоделиСтвола", "Модель ствола", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataNozzleStrui(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String, ByVal keyFieldName2 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "' And " & _
            "[КодВариантаСтвола] =" & GetKeyFieldValue(dbThis, dbOuter, "ВариантыСтволов", "КодВариантаСтвола", "Вариант ствола", rstOuter.Fields("КодВариантаСтвола").Value)

        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "КодВариантаСтвола" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "ВариантыСтволов", "КодВариантаСтвола", "Вариант ствола", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataNozzlePodOuts(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String, ByVal keyFieldName2 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "' And " & _
            "[КодСтруи] =" & GetKeyFieldValue(dbThis, dbOuter, "ТипыСтруй", "КодСтруи", "Вид струи", rstOuter.Fields("КодСтруи").Value)

        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "КодСтруи" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "ТипыСтруй", "КодСтруи", "КодВариантаСтвола", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub

Private Sub ImportDataCommon(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String)
'Прока импорта данных для указанной пользоватлем таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись во внутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "'"
        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                
                If fldOuter.Name = "Баллон" Then
                    If Not IsNull(fldOuter.Value) Then
                        fldThis.Value = GetKeyFieldValue(dbThis, dbOuter, "Баллоны", "КодБаллона", "Обозначение", fldOuter.Value)
                    End If
                Else
                    fldThis.Value = fldOuter.Value
                End If
            Next i
            rstThis.Update
        End If
        
        
        rstOuter.MoveNext
    Loop

    
    
End Sub


Private Sub ImportDataLists(ByRef dbThis As Database, ByRef dbOuter As Database, _
        ByVal tableName As String, ByVal keyFieldName1 As String)
'Прока импорта данных для указанной пользоватлеме таблицы
Dim rstThis As Recordset
Dim rstOuter As Recordset
Dim fldThis As Field
Dim fldOuter As Field
Dim i As Integer
Dim criteria As String
Dim needUpdate As Boolean
    
    Set rstThis = dbThis.OpenRecordset(tableName, dbOpenDynaset)
    Set rstOuter = dbOuter.OpenRecordset(tableName, dbOpenDynaset)
    
    rstOuter.MoveFirst
    Do Until rstOuter.EOF
        'Прверяем имеется ли соотвествующая запись в овнутренней БД
        criteria = keyFieldName1 & "='" & rstOuter.Fields(keyFieldName1) & "'"
        rstThis.FindFirst criteria
        If rstThis.NoMatch Then
            'Если нет - созлаем
            rstThis.AddNew
            needUpdate = True
        Else
            'Если да - Проверяем необходимо ли обновлять запись и если да, обновляем
            needUpdate = NeedToUpdate(rstThis, rstOuter)
            rstThis.Edit
        End If
        
        'Если запись вновь созданная, или требует обновеления, импортируем данные
        If needUpdate Then
            For i = 1 To rstOuter.Fields.Count - 1
                Set fldOuter = rstOuter.Fields(i)
                Set fldThis = rstThis.Fields(fldOuter.Name)
                    fldThis.Value = fldOuter.Value
            Next i
            rstThis.Update
        End If
        
        rstOuter.MoveNext
    Loop


End Sub



Private Function NeedToUpdate(ByRef rstThis As Recordset, ByRef rstOuter As Recordset) As Boolean
'Функция возвращает Истина, если нужно обновлять, и Ложь, если нет
Dim fldChangedThis As Field
Dim fldChangedOuter As Field

    On Error GoTo EX

    Set fldChangedThis = rstThis.Fields("Изменено")
    Set fldChangedOuter = rstOuter.Fields("Изменено")
    
    If fldChangedThis.Value < fldChangedOuter.Value Then
        NeedToUpdate = True
    Else
        NeedToUpdate = False
    End If
    
Exit Function
EX:
    NeedToUpdate = False
End Function

Private Function GetKeyFieldValue(ByRef dbThis As Database, ByRef dbOuter As Database, _
                    ByRef keyTableName As String, ByVal keyFieldName As String, _
                    ByVal textFieldName As String, ByVal textFieldValue As String) As Integer
'Функция возвращает цифровое значение ключа для указанной таблицы и указанного текстового значения
Dim rstTable As Recordset
Dim rstOuterTable As Recordset
Dim textValue As String
    
'    If textFieldValue = "" Then
'        Exit Function
'    End If
    
    Set rstOuterTable = dbOuter.OpenRecordset(keyTableName, dbOpenDynaset)
    rstOuterTable.FindFirst "[" & keyFieldName & "]=" & textFieldValue
    textValue = rstOuterTable.Fields(textFieldName).Value
    
    Set rstTable = dbThis.OpenRecordset(keyTableName, dbOpenDynaset)
    
    rstTable.FindFirst "[" & textFieldName & "]='" & textValue & "'"
'    If rstTable.NoMatch Then
'        rstTable.AddNew
'        rstTable.Fields(textFieldName).Value = textValue
'        rstTable.Update
'    End If
    
    GetKeyFieldValue = rstTable.Fields(keyFieldName).Value
    
Set rstTable = Nothing
End Function
