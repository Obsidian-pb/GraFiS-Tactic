VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_INPPV_CommonData 
   Caption         =   "Общие сведения (ЭСУ ППВ)"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "f_INPPV_CommonData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_INPPV_CommonData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub ShowData(ByVal htmlText As String)
'Прока показывает окно с браузером и загружает содержимое
Dim mDoc As MSHTML.IHTMLDocument
    
    On Error GoTo Tail
    
    htmlText = Replace(htmlText, Asc(34), "'")
    
    'Открываем пустую страницу
    WebBrowser1.Navigate "about:blank"

    Set mDoc = WebBrowser1.Document
    mDoc.Write htmlText
    
    Set mDoc = Nothing
    
    Me.Show
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "ShowData"
End Sub
