Attribute VB_Name = "NewMacros"
Sub Macro1()
    Application.Windows.ItemEx("Конструктор отчетов.vsd:Конструктор:Sheet.3 <ФИГУРА>").Activate

    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 0, visUserValue).RowNameU = "IndexPers"

    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = 121

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Вставить строку")
    Application.ActiveWindow.Shape.AddRow visSectionUser, 0, visTagDefault
    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 0, visUserValue).FormulaForceU = "0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 0, visUserPrompt).FormulaForceU = """"""
    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 1, visUserValue).FormulaForceU = "0"
    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 1, visUserPrompt).FormulaForceU = """"""
    Application.EndUndoScope UndoScopeID1, True

    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 1, visUserValue).RowNameU = "Version"

    Application.ActiveWindow.Shape.CellsSRC(visSectionUser, 1, visUserValue).FormulaU = 1

End Sub
