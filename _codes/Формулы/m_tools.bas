Attribute VB_Name = "m_tools"


Public Function ClearString(ByVal txt As String) As String
Dim tmpVal As Variant
    On Error Resume Next
'    txt = Int(txt)
    txt = Round(txt, 2)
    ClearString = txt
End Function
