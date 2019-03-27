VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Выбор набора"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public IsOK As Boolean
Public SetID As Integer

Private Sub Btn_CLose_Click()
    IsOK = False
    DoCmd.Close
End Sub

Private Sub Btn_OK_Click()
    
    On Error GoTo EX
    
'    IsOK = True
    SetID = Me.CB_Sets.Value
'    Debug.Print SetID
    
    

    DoCmd.Close acForm, "Выбор набора"
    CopyRecordToSet SetID
Exit Sub
EX:
    MsgBox "Необходимо указать набор!", vbCritical
End Sub
