VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommOptForm 
   Caption         =   "Настройки расчетов"
   ClientHeight    =   1500
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   4404
   OleObjectBlob   =   "CommOptForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CommOptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apply_Click()
    A.options.GDZSRezRoundUp = CommOptForm.OptionButton2.Value
    A.Refresh Application.ActivePage.Index
    
    'Обновляем показываемые формы
    RefreshOpenedForms
    
    Me.Hide
End Sub


Public Sub ShowForm()
    CommOptForm.OptionButton2.Value = A.options.GDZSRezRoundUp
    
    Me.Show
End Sub

Public Sub RefreshOpenedForms()
'Обновляем показываемые формы

    If WarningsForm.Visible = True Then WarningsForm.Refresh
    If TacticDataForm.Visible = True Then TacticDataForm.Refresh
    
End Sub

