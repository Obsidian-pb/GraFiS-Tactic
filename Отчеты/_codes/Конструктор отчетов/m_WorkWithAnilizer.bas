Attribute VB_Name = "m_WorkWithAnilizer"
Sub sP_MakeReport()
'Процедура оформления окончательного отчетка из конструктора
'---Объявляем переменные
Dim vsO_EndShape As Visio.Shape
Dim vss_Formula As String


    ActiveWindow.Selection.Group
    Set vsO_EndShape = Application.ActiveWindow.Selection(1)

'---Добавляем раздел и команду меню "Обновить"
    vsO_EndShape.AddSection visSectionAction
    vsO_EndShape.AddNamedRow visSectionAction, "ChangeValue", 0
    vss_Formula = "CALLTHIS(" & Chr(34) & "sP_ChangeValue" & Chr(34) & "," & Chr(34) & "Отчеты" & Chr(34) & ")"
    vsO_EndShape.Cells("Actions.ChangeValue.Action").FormulaU = vss_Formula
    vsO_EndShape.Cells("Actions.ChangeValue.Menu").FormulaU = """Обновить"""
    vsO_EndShape.Cells("Actions.ChangeValue.ButtonFace").FormulaU = 8

'---Добавляем раздел "User" и поля IndexPers и Version
    vsO_EndShape.AddSection visSectionUser
    vsO_EndShape.AddNamedRow visSectionUser, "IndexPers", visTagDefault
    vsO_EndShape.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = 121
    vsO_EndShape.AddNamedRow visSectionUser, "Version", visTagDefault
    vsO_EndShape.CellsSRC(visSectionUser, 1, visUserValue).FormulaU = 1

'---Команду "Обновить" при вбросе отчета
    vsO_EndShape.Cells("EventDrop").FormulaU = vss_Formula
    
'---Применяем слой с индексом 0
    vsO_EndShape.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaForceU = """0"""

Set vsO_EndShape = Nothing
End Sub
