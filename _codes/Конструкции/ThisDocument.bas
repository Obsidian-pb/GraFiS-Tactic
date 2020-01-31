VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 100000

Private ButEvent As c_Buttons


Private Sub app_CellChanged(ByVal Cell As IVCell)
'---Один раз в выполняем обновление иконок на кнопках
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub

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
    
'---Активируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If
    
'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates
End Sub

Private Sub CloseDoc()
'Обрабатываем закрытие документа

'---Деактивируем объект отслеживания нажатия кнопок
    Set ButEvent = Nothing
    
'---Деактивируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then Set app = Nothing
    
'---Удаляем панель управления "СпецФункции"
    RemoveTB_Constructions
End Sub


Private Sub sp_MastersImport()
'---Импортируем мастера

'---Масштаб 1:200
    MasterImportSub "Забор"
    MasterImportSub "Забор2"
    MasterImportSub "Забор3"
    MasterImportSub "Забор4"
    MasterImportSub "ЖДПолотно"
    MasterImportSub "ЖДПолотно2"
    MasterImportSub "Обрыв"
    MasterImportSub "Ров"
    MasterImportSub "Насыпь"
    MasterImportSub "ТрамвайныеПути"
'---Масштаб 1:1000
    MasterImportSub "Забор_1000"
    MasterImportSub "Забор2_1000"
    MasterImportSub "Забор3_1000"
    MasterImportSub "Забор4_1000"
    MasterImportSub "ЖДПолотно_1000"
    MasterImportSub "ЖДПолотно2_1000"
    
End Sub


