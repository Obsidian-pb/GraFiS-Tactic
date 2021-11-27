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

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1





'Private Sub app_SelectionAdded(ByVal Selection As IVSelection)
'    Print Selection.Count
'End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    AddTB
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    RemoveTB
End Sub



Public Sub ActivateApp()
    If app Is Nothing Then
        Set app = Visio.Application
    Else
        Set app = Nothing
    End If
End Sub

Private Sub app_SelectionChanged(ByVal Window As IVWindow)
    Debug.Print Window.Selection.Count
End Sub
