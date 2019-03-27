Attribute VB_Name = "ObjectProperties"
'----------------------------------Модуль изменения свойств объекта--------------------------------


Public Sub sP_ChangeObjectProperties(ShpObj As Visio.Shape)
ShpObj.Delete
'Процедура изменения свойств объекта пожара
'---Объявляем переменные
Dim vpVS_DocShape As Visio.Shape

'---Инициируем объект Шэйп-листа документа
Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---Проверяем имеется ли в секции User Шэйп-листа документа строка City
If vpVS_DocShape.CellExists("User.City", 0) = False Then
    sp_DocumentRowsAdd
End If

'---Открываем окно свойств
PropertiesForm.Show

    
End Sub

Private Sub sp_DocumentRowsAdd()
'Процедура создания строк для свойств объекта пожара
'---Объявляем переменные
Dim vpVS_DocShape As Visio.Shape

    On erro GoTo EX
'---Инициируем объект Шэйп-листа документа
    Set vpVS_DocShape = Application.ActiveDocument.DocumentSheet

'---Проверяем наличие секции User и в случае её отсутствия создаем
    If vpVS_DocShape.SectionExists(visSectionUser, 0) = False Then
        vpVS_DocShape.AddSection (visSectionUser)
    End If

'---Добавляем новые строки
    '---Строка "Населенный пункт"
    vpVS_DocShape.AddNamedRow visSectionUser, "City", visTagDefault
    vpVS_DocShape.Cells("User.City").FormulaU = """Название населенного пункта"""
    vpVS_DocShape.Cells("User.City.Prompt").FormulaU = """Название населенного пункта"""
    '---Строка "Адрес"
    vpVS_DocShape.AddNamedRow visSectionUser, "Adress", visTagDefault
    vpVS_DocShape.Cells("User.Adress").FormulaU = """Адрес объекта пожара"""
    vpVS_DocShape.Cells("User.Adress.Prompt").FormulaU = """Адрес объекта пожара"""
    '---Строка "Объект пожара"
    vpVS_DocShape.AddNamedRow visSectionUser, "Object", visTagDefault
    vpVS_DocShape.Cells("User.Object").FormulaU = """Объект пожара"""
    vpVS_DocShape.Cells("User.Object.Prompt").FormulaU = """Объект пожара"""
    '---Строка "Степень огнестойкости объекта"
    vpVS_DocShape.AddNamedRow visSectionUser, "FireRating", visTagDefault
    vpVS_DocShape.Cells("User.FireRating").FormulaU = """3"""
    vpVS_DocShape.Cells("User.FireRating.Prompt").FormulaU = """Степень огнестойкости"""

Exit Sub
EX:
    SaveLog Err, "sp_DocumentRowsAdd"
End Sub


'Public Function PPP(ShpObj As Visio.Shape) As Boolean
'    ShpObj.Cells("User.Row_1") = 111
'End Function
