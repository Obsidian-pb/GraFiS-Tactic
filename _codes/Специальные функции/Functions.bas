Attribute VB_Name = "Functions"
Sub JPGExportAll(ShpObj As Visio.Shape)
    ShpObj.Delete
    ExportJPG.Show
End Sub




Sub SetAspect(ShpObj As Visio.Shape)
'Процедура устаналвивает новое значение Аспекта для данной страницы
    SetAspect_P
    ShpObj.Delete
End Sub


Public Sub FixZIndex(ShpObj As Visio.Shape)
'Прока исправляет положение фигур ГраФиС относительно стен и прочего
    FixZIndex_P
    ShpObj.Delete
End Sub


'-------------------------Проки для работы из панели--------------------------------
Sub JPGExportAll_P()
    ExportJPG.Show
End Sub

Sub SetAspect_P()
'Процедура устаналвивает новое значение Аспекта для данной страницы
Dim vO_Sheet As Visio.Shape
Dim vs_Aspect As Single

'---Активируем обработчик возварата
    Dim UndoScopeID As Long

Set vO_Sheet = Application.ActiveWindow.Shape

On Error GoTo ex

'---Провряем иеется ли на странице ячейка GFS_Aspect
    If vO_Sheet.CellExists("User.GFS_Aspect", 0) = False Then
    '---Если нет, то создаем со значением 1
        If vO_Sheet.SectionExists(visSectionUser, 0) = False Then 'Проверяем имеется ли секция, еслинет - создаем
            vO_Sheet.AddSection visSectionUser
        End If
            vO_Sheet.AddNamedRow visSectionUser, "GFS_Aspect", visTagDefault
            vO_Sheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---Предлагаем его изменить
    vs_Aspect = _
    CSng(InputBox("Измените значение аспекта по своему желанию. Аспект позаоляет задать дополнительное масштабирование для всех фигур ГраФиС, что может быть удобно при работе в схемах с некорректным масштабом.", _
        "ГраФиС - Настройка аспекта", vO_Sheet.Cells("User.GFS_Aspect").Result(visNumber)))
    
'---Проверяем корректность значения Аспекта
    If vs_Aspect <= 0 Or vs_Aspect > 100 Then
        GoTo ex
    End If

'---Устанавливаем новое значение Аспекта
    vO_Sheet.Cells("User.GFS_Aspect").Formula = vs_Aspect

Set vO_Sheet = Nothing
Application.EndUndoScope UndoScopeID, True

Exit Sub

ex:
MsgBox "Введеное вами значение не может быть установлено в качестве Аспекта! Проверьте правильно ли вы его указали! В качестве значений могут быть использованы только числа от 0,1 до 100!", vbCritical, ThisDocument.Name
Set vO_Sheet = Nothing
Application.EndUndoScope UndoScopeID, True

End Sub


Public Sub FixZIndex_P()
'Прока исправляет положение фигур ГраФиС относительно стен и прочего
Dim vsoSelection As Visio.Selection
    
    On Error GoTo ex
    
    '---перемещаем вперед //Техника;ПТВ;Рукавные линии;Водоисточники;Очаг
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Техника;ПТВ;Рукавные линии;Водоисточники;Очаг")
    Application.ActiveWindow.Selection = vsoSelection
    
    Application.ActiveWindow.Selection.BringToFront
    
    ActiveWindow.DeselectAll
    
    '---перемещаем вперед //ГДЗС;Подписи рукавов
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "ГДЗС;Подписи рукавов;Очаг;Управление СиС")
    Application.ActiveWindow.Selection = vsoSelection
    
    Application.ActiveWindow.Selection.BringToFront
    
    ActiveWindow.DeselectAll
    
    
Exit Sub
ex:
    
End Sub


Public Sub ShapesCountShow()
'Прока показа количества фигур в выборке
Dim vO_ShpItm As Visio.Shape
Dim x, y As Double
    
    On Error GoTo ex
    
    'Определяем на какую фигуру вброшена текущая
    MsgBox "Количество фигур в выделении: " & Application.ActiveWindow.Selection.Count, , "ГраФиС"
    
Exit Sub
ex:
End Sub
