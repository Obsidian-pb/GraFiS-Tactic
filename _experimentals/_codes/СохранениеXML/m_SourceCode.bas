Attribute VB_Name = "m_SourceCode"
'--------------Модуль хранить процедуры для экспорта кода VBA и исходника файла во внешние модули-------------
'------------------Нужен чтобы была возможность коммитить код через ГитХаб------------------
Public Sub SaveSourceCode()

Dim targetPath As String
    
    targetPath = GetCodePath
    ExportVBA targetPath
    ExportXML targetPath
    MsgBox "Исходный код экспортирован"

End Sub

Public Sub ExportVBA(ByVal sDestinationFolder As String)
'Собственно экспорт кода
    Dim oVBComponent As Object
    Dim fullName As String

    For Each oVBComponent In Application.ActiveDocument.VBProject.VBComponents
        If oVBComponent.Type = 1 Then
            ' Standard Module
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 2 Then
            ' Class
            fullName = sDestinationFolder & oVBComponent.Name & ".cls"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 3 Then
            ' Form
            fullName = sDestinationFolder & oVBComponent.Name & ".frm"
            oVBComponent.Export fullName
        ElseIf oVBComponent.Type = 100 Then
            ' Document
            fullName = sDestinationFolder & oVBComponent.Name & ".bas"
            oVBComponent.Export fullName
        Else
            ' UNHANDLED/UNKNOWN COMPONENT TYPE
        End If
        Debug.Print "Сохранен " & fullName
    Next oVBComponent

End Sub

Private Sub ExportXML(ByVal sDestinationFolder As String)
'Сохраняем исходный xml код документа
Dim xmlDoc As DOMDocument60
'Dim xmlNode As IXMLDOMNode
Dim docXMLName As String
Dim docFullName As String

    '---Получаем полное имя документа (включая путь)
    docXMLName = GetDocXMLName
    docFullName = sDestinationFolder & docXMLName
    
    Set xmlDoc = New DOMDocument60
    
    '---Сохраняем документ в формате XML
        '---Создаем копию документа
        Dim visioAppTemp As Visio.Application
        Set visioAppTemp = New Visio.Application
        visioAppTemp.Visible = False
        
        Dim doc As Visio.Document
        Set doc = visioAppTemp.Documents.OpenEx(Application.ActiveDocument.fullName, visOpenCopy)
        doc.SaveAs docFullName
        visioAppTemp.Quit


    '---Получаем содержимое документа и вставляем перевод кареток между тэгами
    xmlDoc.Load docFullName
    xmlDoc.LoadXML Replace(xmlDoc.XML, "><", ">" & Chr(13) & Chr(10) & "<")

    xmlDoc.Save docFullName
End Sub


Private Function GetCodePath() As String
'Возвращает путь к папке с исходными кодами
Dim path As String
Dim docNameWODot As String
    
    '---Путь к текущей папке
    path = Application.ActiveDocument.path
    '---Добавляем название папки с кодами
    path = GetDirPath(path & "_codes")
        
    '---Добавляем путь к папке с кодами ДАННОГО документа
    docNameWODot = Split(Application.ActiveDocument.Name, ".")(0)
    path = GetDirPath(path & "\" & docNameWODot)
    
    GetCodePath = path & "\"
End Function

Private Function GetDirPath(ByVal path As String) As String
'Возвращает путь к папке с указанным именем, если такой папки нет, предварительно создает ее
    '---Проверяем есть ли такая папка, если нет - создаем
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    GetDirPath = path
End Function

Private Function GetDocXMLName() As String
'Возвращаем название документа в формате XML
Dim docAttr() As String
    
    docAttr = Split(Application.ActiveDocument.Name, ".")
    
    Select Case docAttr(1)
        Case Is = "vsd"
            GetDocXMLName = docAttr(0) & ".vdx"
        Case Is = "vss"
            GetDocXMLName = docAttr(0) & ".vsx"
        Case Is = "vst"
            GetDocXMLName = docAttr(0) & ".vtx"
    End Select
End Function


