VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_HTMLDataRedact 
   Caption         =   "Общие данные"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   OleObjectBlob   =   "f_HTMLDataRedact.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_HTMLDataRedact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Форма редактирования HTML данных

Private targetShape As Visio.Shape
'Private viewTemplate As String
'Private dataTemplate As DataTemplateESU



Public Sub FormShow(ByRef shp As Visio.Shape)
    
    On Error GoTo Tail
    
    Set targetShape = shp
    
    Me.tb_HTMLText = shp.Cells("Prop.Common").ResultStr(visUnitsString)
    
    Me.Show
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "FormShow"
End Sub

Private Sub cb_Cancel_Click()
    Me.Hide
End Sub

Private Sub cb_Save_Click()
    
    On Error GoTo Tail
    
    targetShape.Cells("Prop.Common").FormulaU = """" & Me.tb_HTMLText & """"
    
    Me.Hide
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчкиу.", , ThisDocument.Name
    SaveLog Err, "cb_Save_Click"
End Sub
