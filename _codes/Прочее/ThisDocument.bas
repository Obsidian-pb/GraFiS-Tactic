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

Dim WithEvents visApp As Visio.Application
Attribute visApp.VB_VarHelpID = -1



Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    'Удаляем ссылку на приложение
'    Set visApp = Visio.Application
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo Tail
    
'---Обновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---Привязываем объект visApp к ссылке на приложение Visio
'    Set visApp = Visio.Application

'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "Программа вызвала ошибку! Если это будет повторяться, свяжитесь с разработчиком."
End Sub






