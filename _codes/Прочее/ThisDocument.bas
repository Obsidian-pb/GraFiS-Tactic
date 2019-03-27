VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)

    On Error GoTo Tail
    
'---Обновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Проверяем наличие обновлений
    fmsgCheckNewVersion.CheckUpdates

Exit Sub
Tail:
    SaveLog Err, "Document_DocumentOpened"
    MsgBox "Программа вызвала ошибку! Если это будет повторяться, свяжитесь с разработчиком."
End Sub
