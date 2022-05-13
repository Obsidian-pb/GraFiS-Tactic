Attribute VB_Name = "m_Documents"
Option Explicit

Public Function IsDocumentOpened(ByVal docName As String) As Boolean
'Проверяем подключен ли документ с именем docName
Dim doc As Visio.Document
    
    On Error GoTo ex
    
    Set doc = Application.Documents(docName)
    IsDocumentOpened = True
Exit Function
ex:
    IsDocumentOpened = False
End Function
