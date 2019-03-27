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
Private WithEvents ButEvent As Office.CommandBarButton
Attribute ButEvent.VB_VarHelpID = -1



Private Sub ButEvent_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Процедура реагирования на нажатие кнопки "Обновить цветовую схему"
    StyleExport
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'---Создаем панель управления "Цветовые схемы" и добавляем на нее кнопку "Обновить"
    AddTBColorShem
    AddButtonRefresh
    Set ButEvent = Application.CommandBars("Цветовые схемы").Controls("Обновить")
    
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'---Очищаем объект ButEvent и удаляем кнопку "Обновить" с панели управления "Цветовые схемы"
    Set ButEvent = Nothing
    DeleteButtonRefresh
    RemoveTBColorShem
    
End Sub



Public Sub ID_PPP()
    Application.ActiveDocument.Styles.Add InputBox("Укажите название нового стиля"), "", 1, 1, 1
    Debug.Print "Does"
End Sub
