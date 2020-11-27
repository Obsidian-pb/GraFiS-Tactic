Attribute VB_Name = "m_WorkWithHTML"
Option Explicit



Public Sub ShowData(ByRef shp As Visio.Shape, ByVal htmlText As String)
    f_FormulaForm.ShowHTML shp, htmlText
End Sub
