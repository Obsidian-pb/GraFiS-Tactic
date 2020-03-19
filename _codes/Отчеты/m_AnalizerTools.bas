Attribute VB_Name = "m_AnalizerTools"
Option Explicit

Public analizer As InfoCollector2               'Публичный объект анализатора
Private changedNames As Collection

'--------------------------Модуль для хранения инструментальных средств для работы с набором елементов-------------
Public Function A() As InfoCollector2
'Получаем ссылку на объект анализатора
    If analizer Is Nothing Then Set analizer = New InfoCollector2
    Set A = analizer
End Function
Public Sub KillA()
'Уничтожаем объект анализатора
    Set analizer = Nothing
End Sub




Public Sub FindE(ByVal searchString As String, Optional ByRef searchByCallName As Boolean = False)
'Печатаем список ключей элементов с поиском по строке
'Пример: FindE "fire"
Dim elem As Element
Dim coll As Collection
Dim elemShell As ElementsShell
    
    Set elemShell = New ElementsShell
    Set coll = elemShell.GetElementsCollection(searchString, searchByCallName)
    
    For Each elem In coll
        Debug.Print elem.ID
    Next elem
    
End Sub



'-----------------Инструменты обновления сведений в фигурах отчетов---------------------------
Public Sub ReportShapesStringsRefresh() '(ByVal block As Boolean)
'Служебная процедура! Вызывается ТОЛЬКО при разработке функционала трафарета Отчеты
Dim doc As Visio.Document
Dim elements As ElementsShell
Dim CallNames As String
Dim mstr As Visio.Master
Dim shp As Visio.Shape
    
    'Активируем коллекцию принятых ответов
    Set changedNames = New Collection
    
    'Полуаем строку с именами вызова всех элементов
    Set elements = New ElementsShell
    CallNames = elements.CallNames(";")

    'Перебираем все мастера в трафарете и в них перебираем все фигуры.
    Set doc = ThisDocument
    For Each mstr In doc.Masters
        Debug.Print "===========: " & mstr.Name
        For Each shp In mstr.Shapes
            SetPropertyNames shp, CallNames
        Next shp
    Next mstr
    
    On Error GoTo EX
    '!Для экспорта в детали отчетов
    Set doc = Application.Documents("Детали отчетов.vss")
    If Not doc Is Nothing Then
        For Each mstr In doc.Masters
            Debug.Print "===========: " & mstr.Name
            For Each shp In mstr.Shapes
                SetPropertyNames shp, CallNames
            Next shp
        Next mstr
    End If

EX:
    Set changedNames = Nothing
    Set elements = Nothing
End Sub

Private Sub SetPropertyNames(ByRef shp As Visio.Shape, ByVal properyNames As String)
'Если у фигуры имеется ячейка Prop.PropertyName, заменяем ее содержимое новым списком Элементов и пытаемся сохранить имя ранее установленного.
Dim shpChild As Visio.Shape
Dim curCallName As String
Dim newCallName As String
    
    If shp.CellExists("Prop.PropertyName", 0) Then
        curCallName = CellVal(shp, "Prop.PropertyName.Value", visUnitsString)
        If Not InStr(1, properyNames, curCallName, vbTextCompare) > 0 Then
            'Проверяем заменялось ли уже такое имя вызова
            newCallName = GetExistedAnswer(curCallName)
            If newCallName = "" Then
                NewPropSelectForm.OpenForSelect properyNames, curCallName
                If NewPropSelectForm.ok Then
                    shp.Cells("Prop.PropertyName.Format").Formula = """" & properyNames & """"
                    shp.Cells("Prop.PropertyName.Value").Formula = """" & NewPropSelectForm.lbCallNames.Text & """"
                    'Сохраняем принятый вариант для следующих изменений ДОРАБОТАТЬ! ИНОГДА ВЫЗЫВАЕТ ОШИБКУ
                    
'                    changedNames.Add NewPropSelectForm.lbCallNames.Text, curCallName
                End If
            Else
                shp.Cells("Prop.PropertyName.Format").Formula = """" & properyNames & """"
                shp.Cells("Prop.PropertyName.Value").Formula = """" & newCallName & """"
            End If
        End If
    End If
    
    If shp.Shapes.Count > 0 Then
        For Each shpChild In shp.Shapes
            SetPropertyNames shpChild, properyNames
        Next shpChild
    End If
End Sub

Private Function GetExistedAnswer(ByVal curCallName As String) As String
Dim str As Object
    
    On Error GoTo EX

    Set str = changedNames.Item(curCallName)
Exit Function
EX:
    Set str = Nothing
End Function

