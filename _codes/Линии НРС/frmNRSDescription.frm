VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNRSDescription 
   Caption         =   "Îò÷åò î ÍÐÑ"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15135
   OleObjectBlob   =   "frmNRSDescription.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNRSDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function AddNewString(ByVal newStr As String) As frmNRSDescription
    
    Me.tbNRSDescription.Text = Me.tbNRSDescription.Text & _
        newStr & vbNewLine
    
Set AddNewString = Me
End Function
