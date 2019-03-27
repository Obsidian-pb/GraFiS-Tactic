VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Заставка"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Close()
    If Me.Ф_НеПоказывать.Value = True Then
        pS_ChangeShowProperty
    End If

'---Показываем форму "Навигация"
    DoCmd.OpenForm "Навигация", acNormal
End Sub

Private Sub Form_Load()
Dim vss_NotShowSplash, vss_Version, vss_Author As String
'Dim vss_Version As String

'--Загружаем свойства и отображаем на Заставке
'---Загружаем значение свойства "Не показывать заставку"
    vss_NotShowSplash = DFirst("ЗначениеЛогик", "SRVC", "[Описание] = 'SplashNotShow'")
    Me.Ф_НеПоказывать.Value = vss_NotShowSplash
'---Загружаем значение свойства "Версия"
    vss_Version = "v " & DFirst("ЗначениеТекст", "SRVC", "[Описание] = 'Version'")
    Me.П_Версия.Value = vss_Version
'---Загружаем значение свойства "Автор"
'    vss_Author = "Автор идеи и разработчик: " & DFirst("ЗначениеТекст", "SRVC", "[Описание] = 'Author'")
'    Me.Н_АВтор.Caption = vss_Author
    
    
'---Если значение свойства "Не показывать заставку" = True, закрывам заставку
    If vss_NotShowSplash = True Then DoCmd.Close acForm, "Заставка"
        
End Sub

Private Sub Н_Автор_Click()
    DoCmd.Close acForm, "Заставка"
End Sub

Private Sub Надпись14_Click()
    DoCmd.Close acForm, "Заставка"
End Sub

Private Sub Надпись2_Click()
    DoCmd.Close acForm, "Заставка"
End Sub

Private Sub ОбластьДанных_Click()
'Процедура отклика на щелчек левой кнопкой мыши
    DoCmd.Close acForm, "Заставка"
End Sub

Private Sub pS_ChangeShowProperty()
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
        !ЗначениеЛогик = True                      'Изменяет значение на истину
        .Update                                    'Обновляет значение
    End With
    
'---Закрываем все объекты
   vO_rst.Close
   Set vO_dbs = Nothing
End Sub

