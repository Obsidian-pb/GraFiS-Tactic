VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommOptForm 
   Caption         =   "Опции замечаний"
   ClientHeight    =   1500
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4410
   OleObjectBlob   =   "CommOptForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CommOptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iCollector As InfoCollector

Private Sub Apply_Click()
'    CommOptForm.Hide
'    bo_GDZSRezRoundUp = CommOptForm.OptionButton2.Value
'    MasterCheckRefresh
    
'    iCollector.bo_GDZSRezRoundUp
    Me.Hide
    MasterCheckRefresh
End Sub

'Private Sub UserForm_Activate()
'    CommOptForm.OptionButton2.Value = bo_GDZSRezRoundUp
'End Sub

'Public Sub ShowForm(ByRef a_iCollector As InfoCollector)
Public Sub ShowForm()
    Set iCollector = a_iCollector
    
    CommOptForm.OptionButton2.Value = iCollector.bo_GDZSRezRoundUp
    
    Me.Show
End Sub
