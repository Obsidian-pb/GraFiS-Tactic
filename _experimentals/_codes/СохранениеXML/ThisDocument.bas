VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub FDD()
Dim sec As Visio.Section

'    Set sec = Application.ActivePage.PageSheet.Section(visSectionUser)
    Set sec = Application.ActivePage.PageSheet.Section(visSectionCharacter)
    
    
    
    For i = 0 To sec.Count - 1
        Debug.Print sec.Row(i).Name
    Next i
    

End Sub
