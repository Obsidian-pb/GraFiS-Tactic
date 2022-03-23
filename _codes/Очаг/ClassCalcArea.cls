VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCalcArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ComBut As Office.CommandBarButton
Attribute ComBut.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set ComBut = Application.CommandBars("Превращения").Controls("Расчетная зона")
End Sub

Private Sub Class_Terminate()
    Set ComBut = Nothing
End Sub

Private Sub ComBut_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    On Error GoTo ex
    
    ImportCalcAreaInformation
    
Exit Sub
ex:
    SaveLog Err, "CalcArea_ComBut_Click"
End Sub





