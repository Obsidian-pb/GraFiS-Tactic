Attribute VB_Name = "m_WorkWithShapes"
Option Explicit


Public Sub RedactThisText(ByRef shp As Visio.Shape, ByVal cellName As String)
    Debug.Print cellName
    frm_Command.CurrentCommand shp, cellName
End Sub
