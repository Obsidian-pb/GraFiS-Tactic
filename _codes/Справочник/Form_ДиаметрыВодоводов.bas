VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ДиаметрыВодоводов"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(Cancel As Integer)
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub ПСС_ДиаметрВодовода_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DoCmd.Close acForm, "ДиаметрыВодоводов"
        Form_Водоотдача.С_ДиаметрСети.Requery
    End If
End Sub
