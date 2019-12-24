VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Integer
Const cellChangedInterval = 1000

Private ButEvent As c_Buttons


Private Sub app_CellChanged(ByVal Cell As IVCell)
    cellChangedCount = cellChangedCount + 1
    Debug.Print cellChangedCount
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
        Debug.Print "changed"
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
    
'---Активируем объект отслеживания eciaiaiee a i?eei?aiee
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
    
'---Деактивируем объект отслеживания eciaiaiee a i?eei?aiee
    If Application.version > 12 Then Set app = Nothing
    
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


