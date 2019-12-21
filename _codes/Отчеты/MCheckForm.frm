VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MCheckForm 
   Caption         =   " Мастер проверок схемы - Бета-версия"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8628
   OleObjectBlob   =   "MCheckForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As c_chng_shp

Private Sub ListBox1_Click()
Master_check_refresh
End Sub

Private Sub ListBox2_Click()
Master_check_refresh
End Sub

Private Sub UserForm_Activate()
Set c = New c_chng_shp
End Sub

Private Sub UserForm_Click()
Master_check_refresh
End Sub
