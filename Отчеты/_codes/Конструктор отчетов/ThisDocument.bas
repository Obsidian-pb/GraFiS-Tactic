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
'Процедура реагирования на нажатие кнопки "Сделать отчет"
'---Проверяем - выбрана ли группа фигур
    If ActiveWindow.Selection.Count > 1 Then
        sP_MakeReport
    Else
        MsgBox "Не указана группа фигур для добавления в отчет!"
    End If
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'---Создаем панель управления "Отчеты" и добавляем на нее кнопку "Сделать отчет"
    AddTBReport
    AddButtonMakeReport
    Set ButEvent = Application.CommandBars("Отчеты").Controls("Сделать отчет")
    
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
'---Очищаем объект ButEvent и удаляем кнопку "Рукав" с панели управления "Превращения"
    Set ButEvent = Nothing
    DeleteButtonReport
    RemoveTBReport
    
End Sub
