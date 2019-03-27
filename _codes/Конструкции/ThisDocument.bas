VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private ButEvent As c_Buttons

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    OpenDoc
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    CloseDoc
End Sub


Private Sub OpenDoc()
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---Импортируем мастера
    sp_MastersImport
    
'---Создаем панель управления "Конструкции" и добавляем на нее кнопки
    AddTB_Constructions
    AddButtons
    
'---Активируем объект отслеживания нажатия кнопок
    Set ButEvent = New c_Buttons
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates
End Sub

Private Sub CloseDoc()
'Обрабатываем закрытие документа

'---Деактивируем объект отслеживания нажатия кнопок
    Set ButEvent = Nothing
    
'---Удаляем панель управления "СпецФункции"
    RemoveTB_Constructions
End Sub


Private Sub sp_MastersImport()
'---Импортируем мастера

'---Масштаб 1:200
    MasterImportSub "Конструкции.vss", "Забор"
    MasterImportSub "Конструкции.vss", "Забор2"
    MasterImportSub "Конструкции.vss", "Забор3"
    MasterImportSub "Конструкции.vss", "Забор4"
    MasterImportSub "Конструкции.vss", "ЖДПолотно"
    MasterImportSub "Конструкции.vss", "ЖДПолотно2"
    MasterImportSub "Конструкции.vss", "Обрыв"
    MasterImportSub "Конструкции.vss", "Ров"
    MasterImportSub "Конструкции.vss", "Насыпь"
    MasterImportSub "Конструкции.vss", "ТрамвайныеПути"
'---Масштаб 1:1000
    MasterImportSub "Конструкции.vss", "Забор_1000"
    MasterImportSub "Конструкции.vss", "Забор2_1000"
    MasterImportSub "Конструкции.vss", "Забор3_1000"
    MasterImportSub "Конструкции.vss", "Забор4_1000"
    MasterImportSub "Конструкции.vss", "ЖДПолотно_1000"
    MasterImportSub "Конструкции.vss", "ЖДПолотно2_1000"
    
End Sub




'Public Sub TestOpen()
'    OpenDoc
'End Sub
'
'Public Sub TestClose()
'    CloseDoc
'End Sub
