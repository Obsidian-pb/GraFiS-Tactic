VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputDistanceForm 
   Caption         =   "Параметры указателей расстояния"
   ClientHeight    =   1548
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   5160
   OleObjectBlob   =   "InputDistanceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputDistanceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lmax As Integer
Public inppw As Boolean
Public Flag As Boolean   'Флаг указывающий на необходимость обновления данных - если False - не нужно

Private Sub CommandButton1_Click()
On Error Resume Next
    lmax = TextBox1.Value
    inppw = CheckBox1.Value
If Err.Number = 13 Then
    TextBox1.Value = 50
Else
    InputDistanceForm.Hide
    Flag = 1
End If
End Sub

