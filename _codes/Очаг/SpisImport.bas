Attribute VB_Name = "SpisImport"
'------------------------Модуль для процедур импорта списков-------------------
'------------------------Блок независимых списков------------------------------

Public Sub BaseListsRefresh(ShpObj As Visio.Shape)
'Процедура обновления данных фигуры (всех списков)
'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Обновляем список категорий
            ShpObj.Cells("Prop.FireCategorie.Format").FormulaU = ListImport("З_Интенсивности", "Категория")
        
        '---Обновляем список описаний в соответствии со значением категории
            DescriptionsListImport (ShpObj.ID)
        '---Запускаем процедуру получения ЗНАЧЕНИЙ факторов пожара для данного описания
            GetFactorsByDescription (ShpObj.ID)
        
        '---Добавляем ссылку на текущее время страницы
        ShpObj.Cells("Prop.SquareTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If

On Error Resume Next 'НЕ ЗАБЫТЬ ЧТО ВКЛЮЧЕН ОБРАБОТЧИК ОШИБКИ
If VfB_NotShowPropertiesWindow = False Then Application.DoCmd (1312) 'В случае если показ окон включен, показываем окно

End Sub

Public Sub SetRushTime(ShpObj As Visio.Shape)
'Процедура устанавливает время обрушения текущим временем
'---Проверяем вбрасывается ли данная фигура впервые
    If IsFirstDrop(ShpObj) Then
        '---Добавляем ссылку на текущее время страницы
        ShpObj.Cells("Prop.RushTime").Formula = _
            Application.ActiveDocument.DocumentSheet.Cells("User.CurrentTime").Result(visDate)
    End If


End Sub

Public Sub DescriptionsListImport(ShpIndex As Long)
'Процедура импорта списка описаний
'---Объявляем переменные
Dim shp As Visio.Shape
Dim Criteria As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)

'---Устанавливаем значение категории
    Criteria = shp.Cells("Prop.FireCategorie").ResultStr(Visio.visNone)

'---Запускаем процедуру получения относительного списка КАТЕГОРИЙ для текущей фигуры
        If shp.Cells("Prop.IntenseShowType").ResultStr(Visio.visNone) = "По категории" Then
            shp.Cells("Prop.FireDescription.Format").FormulaU = ListImport2("З_Интенсивности", "Описание", "Категория", Criteria)
        End If

'---В случае, если значение поля или формата для нового списка равно "", переводим фокус в ячейке на 0-е положение.
    If shp.Cells("Prop.FireDescription.Format").ResultStr(Visio.visNone) = "" Or shp.Cells("Prop.FireDescription").ResultStr(Visio.visNone) = "" Then
        shp.Cells("Prop.FireDescription").FormulaU = "INDEX(0,Prop.FireDescription.Format)"
    End If

Set shp = Nothing
End Sub











Private Sub ToZeroListIndex(cell As String, ShpIndex As Long) '!!!Временно не используется в связи с отсутствием необхоимости
'---В случае, если значение поля для нового списка равно "", переводим фокус в ячейке на 0-е положение.
Dim CellName As String, CellContent As String

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = Application.ActivePage.Shapes.ItemFromID(ShpIndex)
    
'---Определяем названия ячеек которым будут меняться значения
    CellName = "Prop." & cell
    CellContent = "INDEX(0,Prop." & cell & ".Format)"
    If shp.Cells(CellName).ResultStr(Visio.visNone) = "" Then
        shp.Cells(CellName).FormulaU = CellContent
    End If
End Sub




