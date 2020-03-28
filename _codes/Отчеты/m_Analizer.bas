Attribute VB_Name = "m_Analizer"
Option Explicit
'--------------------------------------------------Модуль для работы с отчетами на странице----------------------------


Public Sub sP_ChangeValue(ShpObj As Visio.Shape)
'Процедура реакции на действие пользователя
Dim targetPage As Visio.Page

'---Предлагаем пользователю указать страницу для анализа
    SeetsSelectForm.Show
    Set targetPage = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet)

'---обновляем имеющиеся значения свойств
    A.Refresh targetPage.Index

'---Запускаем циклы обработки фигур отчета
    sP_ChangeValueMain ShpObj, targetPage

End Sub

Public Sub sP_ChangeValueMain(ByRef vsO_BaseShape As Visio.Shape, ByRef vsO_TargetPage As Visio.Page)
'Процедура реакции на действие пользователя
Dim vsO_Shape As Visio.Shape

    On Error GoTo EX

'---Перебираем все фигуры в отчете
    If vsO_BaseShape.Shapes.Count > 0 Then
        For Each vsO_Shape In vsO_BaseShape.Shapes
            If vsO_Shape.CellExists("Actions.ChangeValue", 0) = True Then '---Проверяем является ли фигура ОТЧЕТОМ!!!
                sP_ChangeValueMain vsO_Shape, vsO_TargetPage
            End If
            If vsO_Shape.CellExists("User.PropertyValue", 0) = True Then '---Проверяем является ли фигура полем отчета
                vsO_Shape.Cells("User.PropertyValue").FormulaU = _
                    str(A.ResultByCN(vsO_Shape.Cells("Prop.PropertyName").ResultStr(visUnitsString)))
            End If
        Next vsO_Shape
    End If

Exit Sub
EX:
    SaveLog Err, "sP_ChangeValueMain"
End Sub

