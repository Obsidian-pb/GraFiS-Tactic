Attribute VB_Name = "m_FormProcs"
Option Compare Database

Public Sub sP_ControlsBlockChange(ass_FormName As String)
'Процедура блокировки-разблокировки элементов управления в формах
Dim vo_Field As control
Dim vo_CurForm As Access.Form

'---Определяем текущую форму для работы с ней
    Set vo_CurForm = Application.Forms(ass_FormName)

'---Перебираем все элементы управления в форме
    For Each vo_Field In vo_CurForm.Controls
        If vo_Field.ControlType = acComboBox Or vo_Field.ControlType = acTextBox _
                Or vo_Field.ControlType = acCheckBox Then       'Если элемент управления является полем или выпадающим списком или флажком
            vo_Field.Locked = Not vo_CurForm.Controls("В_РедактироватьДанные").Value        'Изменяем состояние блокировки элемента в соответствии с состоянием выключателя
        End If
    Next vo_Field
    
End Sub
