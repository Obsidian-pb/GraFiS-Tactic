VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_chng_shp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set app = Visio.Application

End Sub


Private Sub Class_Terminate()
    Set app = Nothing
End Sub

Public Sub app_CellChanged(ByVal Cell As Visio.IVCell)
   MasterCheckRefresh
End Sub

