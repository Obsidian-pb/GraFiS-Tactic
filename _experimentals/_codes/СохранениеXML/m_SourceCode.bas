Attribute VB_Name = "m_SourceCode"
'--------------Модуль хранить процедуры для экспорта кода VBA и исходника файла во внешние модули-------------
'------------------Нужен чтобы была возможность коммитить код через ГитХаб------------------
Public Sub SaveSourceCode()

Dim targetPath As String
    
    targetPath = GetCodePath
    ExportVBA targetPath
    ExportDocState targetPath
    MsgBox "Исходный код экспортирован"

End Sub

Public Sub ExportVBA(ByVal sDestinationFolder As String)
'Собственно экспорт кода
    Dim oVBComponent As Object
    Dim fullName As String

    For Each oVBComponent In Application.ActiveDocument.VBProject.VBComponents
        If oVBComponent.Type = 1 Then
            ' Standard Module
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 2 Then
            ' Class
            fullName = sDestinationFolder & oVBComponent.Name & ".cls"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 3 Then
            ' Form
            fullName = sDestinationFolder & oVBComponent.Name & ".frm"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 100 Then
            ' Document
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        Else
            ' UNHANDLED/UNKNOWN COMPONENT TYPE
        End If
        Debug.Print "Сохранен " & fullName
    Next oVBComponent

End Sub

Private Function GetCodePath() As String
'Возвращает путь к папке с исходными кодами
Dim path As String
Dim docNameWODot As String
    
    '---Путь к текущей папке
    path = Application.ActiveDocument.path
    '---Добавляем название папки с кодами
    path = GetDirPath(path & "_codes")
        
    '---Добавляем путь к папке с кодами ДАННОГО документа
    docNameWODot = Split(Application.ActiveDocument.Name, ".")(0)
    path = GetDirPath(path & "\" & docNameWODot)
    
    GetCodePath = path & "\"
End Function

Private Function GetDirPath(ByVal path As String) As String
'Возвращает путь к папке с указанным именем, если такой папки нет, предварительно создает ее
    '---Проверяем есть ли такая папка, если нет - создаем
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    GetDirPath = path
End Function



'--------------------Работа с состоянием документа (страницы, фигуры, мастера, стили и т.д.)---------
Private Sub ExportDocState(ByVal sDestinationFolder As String)
'Сохраняем состояние документа в текстовый файл
Dim doc As Visio.Document
Dim docFullName As String

Dim pg As Visio.Page
Dim shp As Visio.Shape
Dim mstr As Visio.Master

'---Получаем ссылку на документ и полный путь к нему
    Set doc = Application.ActiveDocument
    docFullName = sDestinationFolder & Replace(doc.Name, ".", "-") & ".txt"
    
'---Очищаем файл, если он уже есть
    With CreateObject("Scripting.FileSystemObject")
        .CreateTextFile docFullName, True
    End With
    
'---Сохраняем состояние всех видов объектов в документе
    '---Документ
    WriteSheetData doc.DocumentSheet, docFullName, "Документ"
    '---Страницы
    For Each pg In doc.Pages
        WriteSheetData pg.PageSheet, docFullName, pg.Name
        'Фигуры
        For Each shp In pg.Shapes
            WriteSheetData shp, docFullName, shp.Name
        Next shp
    Next pg

    '---Мастера
    For Each mstr In doc.Masters
        WriteSheetData mstr.PageSheet, docFullName, mstr.Name
        For Each shp In mstr.Shapes
            WriteSheetData shp, docFullName, shp.Name
        Next shp
    Next mstr
    
    '---Стили
    
    
    '---Узоры заливки
    
    
    '---Шиблоны линий
    
    
    '---Концы линий
    
    
    Debug.Print "Сохранен " & docFullName
End Sub

Public Sub WriteSheetData(ByRef sheet As Visio.Shape, ByVal docFullName As String, ByVal printingName As String)
'Сохраняем в файл по адресу docFullName текущее состояние листа документа, страницы или фигуры (мастера)
Dim shp As Visio.Shape

'---Открываем файл состояния документа
    Open docFullName For Append As #1
    '---Записываем имя объекта
    Print #1, ""
    Print #1, "=>sheet: " & printingName
    '---Закрываем файл состояния документа
    Close #1
    
'---Экспортируем данные по всем возможнымсекциям
    '---Общие
    SaveSectionState sheet, visSectionAction, docFullName
    SaveSectionState sheet, visSectionAnnotation, docFullName
    SaveSectionState sheet, visSectionCharacter, docFullName
    SaveSectionState sheet, visSectionConnectionPts, docFullName
    SaveSectionState sheet, visSectionControls, docFullName
    SaveSectionState sheet, visSectionFirst, docFullName
    SaveSectionState sheet, visSectionFirstComponent, docFullName
    SaveSectionState sheet, visSectionHyperlink, docFullName
    SaveSectionState sheet, visSectionInval, docFullName
    SaveSectionState sheet, visSectionLast, docFullName
    SaveSectionState sheet, visSectionLastComponent, docFullName
    SaveSectionState sheet, visSectionLayer, docFullName
    SaveSectionState sheet, visSectionNone, docFullName
    SaveSectionState sheet, visSectionParagraph, docFullName
    SaveSectionState sheet, visSectionProp, docFullName
    SaveSectionState sheet, visSectionReviewer, docFullName
    SaveSectionState sheet, visSectionScratch, docFullName
    SaveSectionState sheet, visSectionSmartTag, docFullName
    SaveSectionState sheet, visSectionTab, docFullName
    SaveSectionState sheet, visSectionTextField, docFullName
    SaveSectionState sheet, visSectionUser, docFullName
    '---Секция Объект
    SaveSectionObjectState sheet, visRowAlign, docFullName
    SaveSectionObjectState sheet, visRowEvent, docFullName
    SaveSectionObjectState sheet, visRowDoc, docFullName
    SaveSectionObjectState sheet, visRowFill, docFullName
    SaveSectionObjectState sheet, visRowForeign, docFullName
    SaveSectionObjectState sheet, visRowGroup, docFullName
    SaveSectionObjectState sheet, visRowHelpCopyright, docFullName
    SaveSectionObjectState sheet, visRowImage, docFullName
    SaveSectionObjectState sheet, visRowLayerMem, docFullName
    SaveSectionObjectState sheet, visRowLine, docFullName
    SaveSectionObjectState sheet, visRowLock, docFullName
    SaveSectionObjectState sheet, visRowMisc, docFullName
    SaveSectionObjectState sheet, visRowPageLayout, docFullName
    SaveSectionObjectState sheet, visRowPage, docFullName
    SaveSectionObjectState sheet, visRowPrintProperties, docFullName
    SaveSectionObjectState sheet, visRowShapeLayout, docFullName
    SaveSectionObjectState sheet, visRowStyle, docFullName
    SaveSectionObjectState sheet, visRowTextXForm, docFullName
    SaveSectionObjectState sheet, visRowText, docFullName
    SaveSectionObjectState sheet, visRowXForm1D, docFullName
    SaveSectionObjectState sheet, visRowXFormOut, docFullName
    
    
    'Если указанный объект имеет дочерние фигуры - запускаем процедуру сохранения и для них (актуально только для фигур)
    On Error GoTo EX
    If sheet.Shapes.Count > 0 Then
        For Each shp In pg.Shapes
            WriteSheetData shp, docFullName, shp.Name
        Next shp
    End If
    
EX:

End Sub



Private Sub SaveSectionState(ByRef shp As Visio.Shape, ByVal sectID As VisSectionIndices, ByVal docFullName As String)
'Сохраняем в файл по адресу docFullName текущее состояние указанной секции листа документа, страницы или фигуры (мастера)
'ОБЩЕЕ
Dim sect As Visio.Section
Dim rwI As Integer
Dim rw As Visio.Row
Dim cllI As Integer
Dim cll As Visio.Cell
Dim str As String
    
    If shp.SectionExists(sectID, 0) = 0 Then Exit Sub
    Set sect = shp.Section(sectID)
    
    '---Открываем файл состояния документа
    Open docFullName For Append As #1
    '---Записываем индекс Секции
    Print #1, "  Section: " & sectID & ">>>"
    
    '---Перебираем все row секции и для каждой из row формируем строку содержащуюю пары Имя-Формула всех ячеек. При условии, что ячейка не пустая
    For rwI = 0 To sect.Count - 1
        Set rw = sect.Row(rwI)
        str = "    "
        For cllI = 0 To rw.Count - 1
            Set cll = rw.Cell(cllI)
            If cll.Formula <> "" Then
                str = str & cll.Name & ": " & cll.Formula & "; "
            End If
        Next cllI
        'Сохраняем строку в файл
        Print #1, str
    Next rwI
    
    '---Закрываем файл состояния документа
    Close #1
    
Exit Sub
EX:
    Debug.Print "Section ERROR: " & sectID
End Sub

Private Sub SaveSectionObjectState(ByRef shp As Visio.Shape, ByVal rowID As VisRowIndices, ByVal docFullName As String)
'Сохраняем в файл по адресу docFullName текущее состояние листа документа, страницы или фигуры (мастера)
'!!!Для Ячейки ОБЪЕКТ!!!
Dim sect As Visio.Section
Dim rw As Visio.Row
Dim cllI As Integer
Dim cll As Visio.Cell
Dim str As String
    
    If shp.RowExists(visSectionObject, rowID, 0) = 0 Then Exit Sub
    Set sect = shp.Section(visSectionObject)
    
    '---Открываем файл состояния документа
    Open docFullName For Append As #1
    '---Записываем индекс Секции
    Print #1, "  ObjectRow: " & rowID & ">>>"
    
    '---Перебираем все row секции и для каждой из row формируем строку содержащуюю пары Имя-Формула всех ячеек. При условии, что ячейка не пустая
    Set rw = sect.Row(rowID)
    If rw.Count > 0 Then
        str = "    "
        For cllI = 0 To rw.Count - 1
            Set cll = rw.Cell(cllI)
            If cll.Formula <> "" Then
                str = str & cll.Name & ": " & cll.Formula & "; "
            End If
        Next cllI
        'Сохраняем строку в файл
        Print #1, str
    End If
    
    '---Закрываем файл состояния документа
    Close #1
    
Exit Sub
EX:
    Debug.Print "Section Oject ERROR: " & sectID & ", rowID: " & rowID
End Sub
