VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MCheckForm 
   Caption         =   " Мастер проверок схемы"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8670
   OleObjectBlob   =   "MCheckForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub B_Refresh_Click()
    Master_check_refresh
End Sub
