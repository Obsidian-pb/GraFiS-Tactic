VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ПКС"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(Cancel As Integer)
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub Form_Current()

'---Блокируем элементы управления
    В_РедактироватьДанные.Value = False
    В_РедактироватьДанные.Requery
    sP_ControlsBlockChange (Me.Name)
    
End Sub

Private Sub В_РедактироватьДанные_AfterUpdate()
    sP_ControlsBlockChange (Me.Name)
End Sub
