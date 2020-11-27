VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_FormulaForm 
   Caption         =   "Результаты"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10515
   OleObjectBlob   =   "f_FormulaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_FormulaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private state_ As Boolean
Public formShape As Visio.Shape


Public Property Get State() As Boolean
    State = state_
End Property
Public Property Let State(ByVal state_a As Boolean)
    state_ = state_a
    
    If Not state_ Then
        ShowData Me.tb_HTMLCode.Text
        
        wb_Bowser.Visible = True
        tb_HTMLCode.Visible = False
        
        cb_ChangeView.Caption = "Текст"
    Else
        wb_Bowser.Visible = False
        tb_HTMLCode.Visible = True
        
        cb_ChangeView.Caption = "HTML"
    End If
End Property

Private Sub UserForm_Activate()
Dim width As Integer
Dim height As Integer
Dim top As Integer
Dim left As Integer

    
    
    top = 30
    left = 6
    width = Me.InsideWidth - left * 2
    height = 400
    
    

    tb_HTMLCode.width = width
    tb_HTMLCode.height = height
    tb_HTMLCode.top = top
    tb_HTMLCode.left = left
    
    wb_Bowser.width = width
    wb_Bowser.height = height
    wb_Bowser.top = top
    wb_Bowser.left = left
    
    cb_Save.top = top + height + 6
    cb_Cancel.top = top + height + 6
    
    Me.height = cb_Save.top + cb_Save.height + 30
End Sub











Private Sub cb_ChangeView_Click()
    State = Not state_
End Sub

Private Sub cb_Save_Click()
    SaveHTMLPattern
End Sub

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub





Public Sub ShowHTML(ByRef shp As Visio.Shape, ByVal htmlText As String)
    Set formShape = shp
    
    tb_HTMLCode.Text = htmlText
    ShowData htmlText
    Me.Show
End Sub

Public Sub ShowData(ByVal htmlText As String)
'Прока показывает окно с браузером и загружает содержимое
Dim mDoc As MSHTML.IHTMLDocument
    
    On Error GoTo Tail
    
    htmlText = PatternToHTML(htmlText)
    
    'Открываем пустую страницу
    wb_Bowser.Navigate "about:blank"

    Set mDoc = wb_Bowser.Document
    
    mDoc.Write htmlText
    wb_Bowser.Refresh
       
    Set mDoc = Nothing
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
'    SaveLog Err, "ShowData"
End Sub


Private Sub SaveHTMLPattern()
    SetCellVal formShape, "User.TextPattern", tb_HTMLCode.Text
End Sub




Private Function PatternToHTML(ByVal htmlText As String) As String
'Функция чистит код исходника HTML и заменяет вставленные ссылки на ячейки фигуры фактическими их значениями
Dim cellVal As String
Dim cll As Visio.Cell
Dim i As Integer
'dim cellName as String
    
    htmlText = Replace(htmlText, Asc(34), "'")
    
    For i = 0 To formShape.RowCount(visSectionProp) - 1
        Set cll = formShape.CellsSRC(visSectionProp, i, visCustPropsValue)
        htmlText = Replace(htmlText, "$" & cll.RowNameU & "$", ClearString(cll.ResultStr(visUnitsString)))
    Next i

PatternToHTML = htmlText
End Function




'<p><b>sdfgsdg</b></p>
'<div>Площадь = $Prop.Square$</div>






