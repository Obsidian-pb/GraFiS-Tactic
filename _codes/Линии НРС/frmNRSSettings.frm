VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNRSSettings 
   Caption         =   "Настройки расчета НРС"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "frmNRSSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNRSSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnSave_Click()
    SaveData
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    FillData
End Sub



Private Sub sbRoundAccuracy_Change()
    Me.tbRoundAccuracy = Me.sbRoundAccuracy.Value
End Sub
Private Sub sbCheckAccuracy_Change()
    Me.tbCheckAccuracy = Me.sbCheckAccuracy.Value
End Sub
Private Sub sbOutAccuracy_Change()
    Me.tbOutAccuracy = Me.sbOutAccuracy.Value
End Sub
Private Sub sbMaxIterations_Change()
    Me.tbMaxIterations = Me.sbMaxIterations.Value
End Sub
Private Sub sbApprovedHout_Change()
    Me.tbApprovedHout = Me.sbApprovedHout.Value
End Sub





Private Sub FillData()
    Me.tbRoundAccuracy = GetSetting("GraFiS", "GFS_NRS", "RoundAccuracy", 4)
    Me.tbCheckAccuracy = GetSetting("GraFiS", "GFS_NRS", "CheckAccuracy", 2)
    Me.tbOutAccuracy = GetSetting("GraFiS", "GFS_NRS", "OutAccuracy", 2)
    Me.tbMaxIterations = GetSetting("GraFiS", "GFS_NRS", "MaxIterations", 100)
    Me.tbApprovedHout = GetSetting("GraFiS", "GFS_NRS", "ApprovedHout", 3)
    
    Me.sbRoundAccuracy = Me.tbRoundAccuracy
    Me.sbCheckAccuracy = Me.tbCheckAccuracy
    Me.sbOutAccuracy = Me.tbOutAccuracy
    Me.sbMaxIterations = Me.tbMaxIterations
    Me.sbApprovedHout = Me.tbApprovedHout
End Sub

Private Sub SaveData()
    SaveSetting "GraFiS", "GFS_NRS", "RoundAccuracy", Me.tbRoundAccuracy
    SaveSetting "GraFiS", "GFS_NRS", "CheckAccuracy", Me.tbCheckAccuracy
    SaveSetting "GraFiS", "GFS_NRS", "OutAccuracy", Me.tbOutAccuracy
    SaveSetting "GraFiS", "GFS_NRS", "MaxIterations", Me.tbMaxIterations
    SaveSetting "GraFiS", "GFS_NRS", "ApprovedHout", Me.tbApprovedHout
End Sub

