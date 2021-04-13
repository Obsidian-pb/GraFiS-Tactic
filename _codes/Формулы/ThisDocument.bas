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


Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    'Проверяем, видим ли трафыарет "Отчеты"
    CheckReportsStencil
End Sub



Public Sub CheckReportsStencil()
'Проверяем подключен ли уже трафарет "Отчеты"
Const rep = "Отчеты.vss"
Dim stenc As Visio.Document
    
    For Each stenc In Application.Documents
        If stenc.name = rep Then
            'stenc.
            Exit Sub
        End If
    Next stenc
    
    Application.Documents.Open (ThisDocument.path & rep)
End Sub


