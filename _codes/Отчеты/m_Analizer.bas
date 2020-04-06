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


'------------------------Для радиальной диаграммы-----------------------------------------
Public Sub sP_ChangeRDValue(ShpObj As Visio.Shape)
'Процедура запуска обновления данных в радиальной диаграмме
Dim targetPage As Visio.Page
'Dim tmp As Variant

'---Предлагаем пользователю указать страницу для анализа
    SeetsSelectForm.Show
    Set targetPage = Application.ActiveDocument.Pages(SeetsSelectForm.SelectedSheet)

'---обновляем имеющиеся значения свойств
    A.Refresh targetPage.Index

'---Запускаем устанавливаем данные для полей радиальной диаграммы
    '---Личного состава
    ShpObj.Cells("Prop.Personnel").Formula = """" & A.Result("PersonnelHave") & "/" & A.Result("PersonnelNeed") & """"
    '---Основных ПА
    ShpObj.Cells("Prop.MainPA").Formula = """" & A.Result("MainPAHave") & "/" & A.Result("ACNeed") & """"
    '---Расход воды
    ShpObj.Cells("Prop.WaterExpense").Formula = """" & A.Result("FactStreamW") & "/" & A.Result("NeedStreamW") & """"
    '---Запас воды (Если воды бесконечное количство, то значение имеющегося запаса воды должно указываться равным фактическому)
'    tmp = A.Result("WaterValueNeed10min")
    If A.Result("WaterEternal") Then
        ShpObj.Cells("Prop.WaterValue").Formula = """" & A.Result("WaterValueHave") & "/" & A.Result("WaterValueHave") & """"
    Else
        ShpObj.Cells("Prop.WaterValue").Formula = """" & A.Result("WaterValueHave") & "/" & A.Result("WaterValueNeed10min") & """"
    End If
    '---Звеньев ГДЗС
    ShpObj.Cells("Prop.GDZS").Formula = """" & A.Sum("GDZSChainsCountWork;GDZSChainsRezCountHave") & "/" & A.Result("GDZSChainsCountNeed") & """"
    '---Стволов
    ShpObj.Cells("Prop.Stv").Formula = """" & A.Result("StvolWHave") & "/" & PF_RoundUp(A.Result("NeedStreamW") / ShpObj.Cells("Prop.StvWaterExpense").Result(visNumber)) & """"

End Sub
