VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Водоотдача"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Current()

'---Блокируем элементы управления
    В_РедактироватьДанные.Value = False
    В_РедактироватьДанные.Requery
'    sP_ControlsBlockChange (Me.Name)
    ПФ_Водоотдача.Locked = Not Me.Controls("В_РедактироватьДанные").Value
    К_ДобавитьДиаметр.Enabled = Me.Controls("В_РедактироватьДанные").Value
    К_РедактироватьДиаметр.Enabled = Me.Controls("В_РедактироватьДанные").Value
    
End Sub

Private Sub В_РедактироватьДанные_AfterUpdate()
'    sP_ControlsBlockChange (Me.Name)
    ПФ_Водоотдача.Locked = Not Me.Controls("В_РедактироватьДанные").Value
    К_ДобавитьДиаметр.Enabled = Me.Controls("В_РедактироватьДанные").Value
    К_РедактироватьДиаметр.Enabled = Me.Controls("В_РедактироватьДанные").Value
    
End Sub

Private Sub С_ВидСети_AfterUpdate()
    С_ДиаметрСети.Requery
    С_ДиаметрСети.Value = DFirst("КодДиаметра", "[Диаметры водоводов]", "КодВидаСети = [С_ВидСети] ") 'Подозрительно
    ПФ_Водоотдача.Requery
End Sub

Private Sub С_ДиаметрСети_AfterUpdate()
    ПФ_Водоотдача.Requery
End Sub
