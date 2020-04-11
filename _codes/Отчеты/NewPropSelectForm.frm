VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewPropSelectForm 
   Caption         =   "Выберите новое название свойства"
   ClientHeight    =   4428
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7884
   OleObjectBlob   =   "NewPropSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewPropSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ok As Boolean


Private Sub cbCancel_Click()
    ok = False
    Me.Hide
End Sub

Public Sub OpenForSelect(ByVal CallNames As String, ByVal prevCallName As String)
'Открываем форму для выбора нового имени вызова для элемента расчета
Dim callNamesArr() As String
Dim i As Integer
    
    Me.tbPrevCallName.Text = prevCallName
    
    'Заполняем варианты новых имен вызова
    callNamesArr = Split(CallNames, ";")
    Me.lbCallNames.Clear
    For i = 0 To UBound(callNamesArr)
        Me.lbCallNames.AddItem callNamesArr(i)
    Next i
    
    ok = False
    Me.Show
End Sub

Private Sub cbOK_Click()
    ok = True
    Me.Hide
End Sub
