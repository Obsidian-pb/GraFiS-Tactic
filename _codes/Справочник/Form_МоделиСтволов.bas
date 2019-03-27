VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_МоделиСтволов"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_BeforeUpdate(Cancel As Integer)
'On Error Resume Next
'    TB_LastChangedTime.Value = Now()
'Dim rst As Recordset

'    Set rst = Me.Recordset
'    rst.Edit
'        rst.Fields("Изменено").Value = Now()
'    rst.Update

End Sub

Private Sub Form_Current()
'Процедура перехода между записями
'---Блокируем элементы управления
    В_РедактироватьРасходы.Value = False
    s_ControlsBlockChange
    
'---Предварительное задание свойств элементов управления
'    Me.К_ОткрытьФормуВариантов.Enabled = False

'---Обновляем списки Вариантов и Струй стволов
    Me.С_КодВариантаСтвола.Requery
    Me.С_КодВариантаСтвола.Value = Me.С_КодВариантаСтвола.ItemData(0)
    Me.ПФ_СтруиСтволов.Requery
    
'---Присваиваем полю Кратность значение из списка С_КодВариантаСтвола
    Me.П_Кратность = Me.С_КодВариантаСтвола.Column(3)
    
End Sub

Private Sub Form_Load()
    Me.ПСС_ТипыСтволов.Enabled = False
    Me.ПСС_ТипыСтволов.Value = 1
    Me.Ф_ОтборПоТипуСтволов.Value = False
    Me.ПСС_КМодели.Value = " "
    Me.ПСС_КМодели.Requery
End Sub

Private Sub В_РедактироватьРасходы_AfterUpdate()
    s_ControlsBlockChange
    TB_LastChangedTime.Value = Now()
End Sub

Private Sub К_ОткрытьФормуВариантов_Click()
    DoCmd.OpenForm "ВариантыСтволов", acNormal
End Sub

Private Sub П_Кратность_AfterUpdate()
'Процедура обновления значения поля Кратность в таблице ВариантыСтволов
'---Проверяем указан ли для текущего ствола вариант и если нет выдаем предупреждение и выходим
    If Me.С_КодВариантаСтвола.ListCount = 0 Then
        MsgBox "Сначала необходимо указать хоть один вариант ствола!", vbInformation
        П_Кратность = ""
        Exit Sub
    End If
    
'---Объявляем переменные
    Dim rst As DAO.Recordset, dbs As DAO.Database
    
'---Присваиваем переменным объекты
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("ВариантыСтволов")
    
'---Ищем запись в таблице Вариатны стволов соответствующую коду указанному в первой колонке списка С_КодВариантаСтвола_
'    И присваиваем ей значение введенное пользователем в поле П_Кратность
    With rst
        .FindFirst ("КодВариантаСтвола =" & Me.С_КодВариантаСтвола.Column(0))
        .Edit
        ![Кратность] = П_Кратность.Value
        .Update
    End With
    
    rst.Close
    dbs.Close
    
    Me.С_КодВариантаСтвола.Requery
End Sub

Private Sub ПСС_КМодели_AfterUpdate()
    ' Поиск записи, соответствующей этому элементу управления.
    Dim rs As Object

'    On Error GoTo EX

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[КодМоделиСтвола] = " & str(Nz(Me.ПСС_КМодели.Value, 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
'EX:
End Sub

Private Sub ПСС_ТипыСтволов_AfterUpdate()
    Me.ПСС_КМодели.Requery
    Me.Requery
End Sub


Private Sub С_КодВариантаСтвола_AfterUpdate()
    Me.ПФ_СтруиСтволов.Requery
    Me.П_Кратность = Me.С_КодВариантаСтвола.Column(3)
End Sub

Private Sub s_ControlsBlockChange()
'Процедура блокировки-разблокировки элементов управления в форме МоделиСтволов
'---Для элементов формы
    Me.П_МодельСтвола.Locked = Not В_РедактироватьРасходы.Value
    Me.ПСС_КодТипаСтвола.Locked = Not В_РедактироватьРасходы.Value
    Me.П_УсловныйПроход.Locked = Not В_РедактироватьРасходы.Value
    Me.П_Кратность.Locked = Not В_РедактироватьРасходы.Value
    Me.К_ОткрытьФормуВариантов.Enabled = В_РедактироватьРасходы.Value
    
'---Для подчиненных форм
    ПФ_СтруиСтволов.Locked = Not В_РедактироватьРасходы.Value
    Me.ПФ_СтруиСтволов.Controls(2).Locked = Not В_РедактироватьРасходы.Value
    
'---Для списка вариантов стволов: возможность внесения изменений
    If В_РедактироватьРасходы.Value = True Then
        Me.С_КодВариантаСтвола.ListItemsEditForm = "ВариантыСтволов"
    Else
        Me.С_КодВариантаСтвола.ListItemsEditForm = ""
    End If
    
'---Для ПФ_СтруиСтвола
    Me.ПФ_СтруиСтволов.Controls("Вид струи").Locked = Not В_РедактироватьРасходы.Value
End Sub


Private Sub Ф_ОтборПоТипуСтволов_AfterUpdate()
    Me.ПСС_ТипыСтволов.Enabled = Ф_ОтборПоТипуСтволов
    Me.ПСС_КМодели.Requery
    Me.Requery
End Sub


