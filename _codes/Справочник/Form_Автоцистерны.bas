VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Автоцистерны"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Btn_GotoWF_Click()
    GotoWF Nz(Ссылка_WF.Value, ""), Nz(Me.Модель.Value, "")
End Sub


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



'-----------------------------------------------------------------------------------------------------------------------
'НА БУДУЩЕЕ:
'Private Sub Form_Current()
'---Процедура для определения относительного запаса воды в цистерне относительно максимального значения по набору данных
'    П_Прог_Воды.Width = Me.ЗапасВоды.Width * (Nz(Me.ЗапасВоды.Value, 0) / DMax("[Запас воды]", "Автоцистерны"))
'End Sub
