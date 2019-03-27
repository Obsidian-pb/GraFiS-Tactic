VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ВариантыСтволов"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_BeforeUpdate(Cancel As Integer)
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub Form_Close()
    Form_МоделиСтволов.С_КодВариантаСтвола.Requery
    Form_МоделиСтволов.С_КодВариантаСтвола.Value = Form_МоделиСтволов.С_КодВариантаСтвола.ItemData(0)
End Sub
