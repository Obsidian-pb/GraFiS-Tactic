VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ПФ_ПроизводиетльностиСтволов"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(Cancel As Integer)
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub Производительность_пены_AfterUpdate()

    PodRefresh

End Sub

Private Sub Расход_AfterUpdate()

    If Form_ПФ_СтруиСтволов.Вид_струи.Column(1) = "Пенная" Then
        Me.Производительность_пены = Me.Расход.Value * Form_МоделиСтволов.П_Кратность.Value
        PodRefresh
    Else
        Me.Производительность_пены = ""
        PodRefresh
    End If

End Sub

Private Sub PodRefresh()
    If Производительность_пены.Value <> "" Then
        Me.Расход_воды_1.Value = Расход.Value * 0.94
        Me.Расход_воды_2.Value = Расход.Value * 0.96
        Me.Расход_ПО_1.Value = Расход.Value * 0.06
        Me.Расход_ПО_2.Value = Расход.Value * 0.04
    Else
        Me.Расход_воды_1.Value = ""
        Me.Расход_воды_2.Value = ""
        Me.Расход_ПО_1.Value = ""
        Me.Расход_ПО_2.Value = ""
    End If
End Sub
