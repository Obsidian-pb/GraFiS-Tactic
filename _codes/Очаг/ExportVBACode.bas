Attribute VB_Name = "ExportVBACode"
'--------------Модуль хранить процедуры для экспорта кода VBA во внешние модули-------------
'------------------Нужен чтобы была возможность коммитить код через ГитХаб------------------
Dim doc As Visio.Document

Public Sub SaveVBACode()

    Set doc = Application.Documents("Очаг.vss")

    ExportVBA doc.path & "\Моделирование\"
    MsgBox "VBA код экспортирован"

End Sub

Public Sub ExportVBA(sDestinationFolder As String)
'Собственно экспорт кода
    Dim oVBComponent As Object
    Dim fullName As String
    
    For Each oVBComponent In doc.VBProject.VBComponents
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


