VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeetsSelectForm 
   Caption         =   "Выбор страницы"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   OleObjectBlob   =   "SeetsSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SeetsSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedSheet As String

Private Sub B_OK_Click()
    
    If OB_CurrentSheet.Value = True Then
        SelectedSheet = Application.ActivePage.Name
    Else
        SelectedSheet = LB_SheetsList.Value
    End If
'    MsgBox SelectedSheet
    
    Me.Hide
End Sub

Private Sub OB_CurrentSheet_Change()
    LB_SheetsList.Locked = True
    LB_SheetsList.BackColor = &H8000000F
    LB_SheetsList.Clear
End Sub

Private Sub OB_SelectSheet_Change()
    LB_SheetsList.Locked = False
    LB_SheetsList.BackColor = vbWhite
    s_SheetsListCreate
End Sub



Private Sub UserForm_Activate()
    If OB_CurrentSheet.Value = False Then s_SheetsListCreate
End Sub

Private Sub s_SheetsListCreate()
'Процедура формирования списка страниц
Dim i As Integer

    LB_SheetsList.Clear
    For i = 1 To Application.ActiveDocument.Pages.Count
        LB_SheetsList.AddItem Application.ActiveDocument.Pages(i).Name
    Next i
    LB_SheetsList.ListIndex = 0

End Sub
