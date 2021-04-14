VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_FormulaFormAll 
   Caption         =   "Результаты"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   OleObjectBlob   =   "f_FormulaFormAll.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_FormulaFormAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
Dim width As Integer
Dim height As Integer
Dim top As Integer
Dim left As Integer

    'Устанавливаем размеры формы и содержимого
    top = 6
    left = 6
    width = Me.InsideWidth - left * 2
    height = 400
    
    wb_Bowser.width = width
    wb_Bowser.height = height
    wb_Bowser.top = top
    wb_Bowser.left = left
    
    cb_Cancel.top = top + height + 6
    
    Me.height = cb_Cancel.top + cb_Cancel.height + 30
    
End Sub


Public Sub ShowData(ByVal htmlText As String)
'Прока показывает окно с браузером и загружает содержимое
Dim mDoc As MSHTML.IHTMLDocument
    
    On Error GoTo Tail
    
    'Открываем пустую страницу
    wb_Bowser.Navigate "about:blank"

    Set mDoc = wb_Bowser.Document
    
    mDoc.Write htmlText
    wb_Bowser.Refresh
       
    Set mDoc = Nothing
    
    Me.Show
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.name
    SaveLog Err, "ShowData", "Формулы - f_formulaFormAll"
End Sub
