Attribute VB_Name = "m_BaseProcs"
Option Compare Database


Public Sub PS_SplashFormShow(ByVal control As IRibbonControl)
'Процедура показа заставки - для кнопки на ленте
'---Изменяем служебное свойство базы "Показывать заставку" на Ложь
    pS_ChangeShowPropertyFalse
    
'---Показываем заставку
    DoCmd.OpenForm "Заставка", acNormal

End Sub

Public Sub PS_NavFormShow(ByVal control As IRibbonControl)
'Процедура показа формы навигации - для кнопки на ленте
'---Показываем форму навигации
    DoCmd.OpenForm "Навигация", acNormal

End Sub

Public Sub PS_Import(ByVal control As IRibbonControl)
'Процедура импорта данных - для кнопки на ленте
'---Запускаем проку импорта данных
    BaseDataImport

End Sub

Public Sub PS_HelpShow(ByVal control As IRibbonControl)
'Процедура показа справки - для кнопки на ленте
'---Показываем справку
    Dim f_pth As String
    
    f_pth = "hh.exe " & Application.CurrentProject.path & "\" & "ГраФиС-Справка.chm"
    Shell f_pth, vbNormalFocus

End Sub

Public Sub PS_CopyToAnotherSet(ByVal control As IRibbonControl)
'Процедура копирования записи в другой набор - для кнопки на ленте
'---Показываем справку
    Dim f_pth As String
    
'    CopyRecordToSet
    DoCmd.OpenForm "Выбор набора"
End Sub




'---------------------------------------Служебные процедуры--------------------------------------------------
Private Sub pS_ChangeShowPropertyFalse()
'Процедура изменения значения служебного свойства базы "Показывать заставку"
'---ОБъявляем переменные
Dim vO_dbs As Database
Dim vO_rst As Recordset


'---Определяем объекты процедуры
    Set vO_dbs = CurrentDb
    Set vO_rst = vO_dbs.OpenRecordset("SRVC", dbOpenDynaset)
    
'---Находим необходимую запись и изменяем её
    With vO_rst
        .FindFirst "[Описание] = 'SplashNotShow'"
        .Edit                                      'Разрешает редактирование.
        !ЗначениеЛогик = False                      'Изменяет значение на истину
        .Update                                    'Обновляет значение
    End With
    
'---Закрываем все объекты
   vO_rst.Close
   Set vO_dbs = Nothing
End Sub

Public Sub CopyRecordToSet(ByVal ttxSet As Integer)
'Прока копирует данные текущей записи в форме в новый набор
Dim frm As Access.Form
Dim rst As Recordset
Dim recordData() As Variant
Dim fld As Field
Dim i As Integer
   
    Set frm = Application.Screen.ActiveForm
    Set rst = frm.Recordset
    
    ReDim recordData(rst.Fields.Count - 1)
    
'---Обновляем перечень наборов
    frm.Controls("ПСС_Набор").Requery
    
'---Формируем набор данных для последующего кописрования
    For i = 1 To rst.Fields.Count - 1
        Set fld = rst.Fields(i)
        recordData(i - 1) = fld.Value
    Next i
    
'---Создаем новую запись и наполняем ее данными
    rst.AddNew
    For i = 0 To UBound(recordData) - 1
        Set fld = rst.Fields(i + 1)
        If fld.Name = "Набор" Then
            fld.Value = ttxSet
        Else
            fld.Value = recordData(i)
        End If
        
    Next i
    rst.Update
    frm.Refresh
    
'---Даем пользователю возможность вносить дальнейшие изменения
    frm.Controls("В_РедактироватьДанные").Value = True
    sP_ControlsBlockChange (frm.Name)
    
End Sub
