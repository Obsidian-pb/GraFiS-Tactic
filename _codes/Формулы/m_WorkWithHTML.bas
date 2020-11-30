Attribute VB_Name = "m_WorkWithHTML"
Option Explicit



Public Sub ShowData(ByRef shp As Visio.Shape, ByVal htmlText As String)
    f_FormulaForm.ShowHTML shp, htmlText
End Sub

Public Sub ShowDataInShape(ByRef shp As Visio.Shape, ByVal htmlText As String)
    f_FormulaForm.CopyBrowserContent shp, htmlText
    SetShowDataInShapeControlFormula shp, "User.DataChangeAction.Prompt"
End Sub

Private Sub SetShowDataInShapeControlFormula(ByRef shp As Visio.Shape, ByVal cellName As String)
Dim i As Integer
Dim frml As String
Dim nameOfRow As String
    
    frml = ""
    For i = 0 To shp.RowCount(visSectionProp) - 1
        nameOfRow = shp.CellsSRC(visSectionProp, i, visCustPropsValue).RowNameU
        If Len(frml) > 0 Then frml = frml & "&"
        frml = frml & "Prop." & nameOfRow
    Next i
    
    SetCellFrml shp, cellName, frml
End Sub

Public Sub ClearShapeText(ByRef shp As Visio.Shape)
    shp.Text = ""
End Sub
