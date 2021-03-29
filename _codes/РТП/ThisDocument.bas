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
    AddTBImagination
    AddButtons
End Sub


Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    DeleteButtons
    RemoveTBImagination
End Sub
